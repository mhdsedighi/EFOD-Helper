import os
import win32com.client as win32
import pandas as pd
import re
from openpyxl import load_workbook
from openpyxl.worksheet.table import Table, TableStyleInfo
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext
import logging
import xml.etree.ElementTree as ET

# Custom logging handler to output to a Tkinter Text widget
class TextHandler(logging.Handler):
    def __init__(self, text_widget, root):
        super().__init__()
        self.text_widget = text_widget
        self.root = root  # Reference to Tkinter root for updating

    def emit(self, record):
        msg = self.format(record)
        self.text_widget.configure(state='normal')
        self.text_widget.insert(tk.END, msg + '\n')
        self.text_widget.configure(state='disabled')
        self.text_widget.see(tk.END)  # Auto-scroll to the bottom
        self.root.update()  # Force GUI update to show message immediately

def setup_logging(text_widget, root):
    # Set up logging to output to the text widget
    logging.basicConfig(level=logging.DEBUG, format='%(asctime)s - %(levelname)s - %(message)s')
    handler = TextHandler(text_widget, root)
    handler.setFormatter(logging.Formatter('%(asctime)s - %(levelname)s - %(message)s'))
    logging.getLogger().handlers = []  # Clear default handlers
    logging.getLogger().addHandler(handler)

def export_table_to_excel(file_path, output_dir, root):
    if not os.path.exists(file_path):
        logging.error(f"File not found: {file_path}")
        return None

    # Initialize Word application
    try:
        word = win32.Dispatch('Word.Application')
        word.Visible = False
        word.DisplayAlerts = False
        logging.info("Word application initialized")
        root.update()  # Update GUI
    except Exception as e:
        logging.error(f"Failed to initialize Word: {e}")
        return None

    try:
        # Open the document
        doc = word.Documents.Open(os.path.abspath(file_path))
        logging.info(f"Opened document: {file_path}")
        root.update()

        # Define Word constants
        WD_NO_PROTECTION = -1

        # Handle protection
        protection_password = None  # Set to your password if known
        if doc.ProtectionType != WD_NO_PROTECTION:
            logging.info("Document is protected. Attempting to unprotect...")
            try:
                if protection_password:
                    doc.Unprotect(protection_password)
                else:
                    doc.Unprotect()
                logging.info("Document unprotected successfully.")
                root.update()
            except Exception as e:
                logging.error(f"Failed to unprotect document: {e}")
                doc.Close()
                return None

        if doc.Tables.Count == 0:
            logging.error("No tables found in the document.")
            doc.Close()
            return None

        # Get the first table
        table = doc.Tables(1)

        # Verify Word table has 11 columns
        if table.Columns.Count != 11:
            logging.error(f"Expected 11 columns in Word table, found {table.Columns.Count}")
            doc.Close()
            return None

        # Prepare data structure
        table_data = []
        checkbox_columns = range(4, 10)  # Columns 4-9 (1-based: 4, 5, 6, 7, 8, 9)

        # Fixed checkbox text mapping
        checkbox_text_map = {
            '4': 'No Difference',
            '5': 'More Exacting',
            '6': 'Different in character',
            '7': 'Less protective or partially',
            '8': 'Significant Difference',
            '9': 'Not Applicable'
        }

        # Iterate through rows (1-based indexing)
        max_rows = table.Rows.Count
        # max_rows = min(30, table.Rows.Count)  # Limit to first 30 rows; comment out to process all rows
        logging.info(f"Processing up to {max_rows} rows")
        root.update()
        for row_idx in range(1, max_rows + 1):  # Process up to max_rows
            row_data = []
            checked_indices = []

            # Iterate through all 11 columns
            for col_idx in range(1, 12):  # 1-based: 1 to 11
                cell = table.Cell(row_idx, col_idx)
                # Get raw text
                raw_text = cell.Range.Text
                # Split at first control character to get visible text
                visible_text = ''
                for char in raw_text:
                    if ord(char) < 32 and char != '\n':
                        break
                    visible_text += char
                visible_text = visible_text.strip()
                # If in first column, extract numeric part from visible text
                if col_idx == 1:
                    numeric_match = re.match(r'^\d+(?:\.\d+)?(?![.\d])', visible_text)
                    if numeric_match:
                        cell_text = numeric_match.group(0)
                        logging.debug(
                            f"Row {row_idx}, Col {col_idx} raw: {repr(raw_text)}, visible: {repr(visible_text)}, matched: {cell_text}, remainder: {repr(raw_text[len(visible_text):].strip())}")
                    else:
                        cell_text = visible_text
                        logging.debug(
                            f"Row {row_idx}, Col {col_idx} raw: {repr(raw_text)}, visible: {repr(visible_text)}, cleaned: {cell_text} (no numeric match)")
                else:
                    cell_text = visible_text

                # Handle checkbox columns (4-9)
                if col_idx in checkbox_columns:
                    for field in cell.Range.FormFields:
                        if field.Type == 71:  # wdFieldFormCheckBox
                            if field.CheckBox.Value:
                                checked_indices.append(str(col_idx))
                            # Debug: Print raw cell content if problematic
                            if any(ord(c) < 32 and c != '\n' for c in raw_text):
                                logging.debug(f"Row {row_idx}, Col {col_idx} raw content: {repr(raw_text)}")
                # Add specific columns to row_data in desired order
                elif col_idx in [1, 2, 3, 10, 11]:  # Only these go to Excel
                    if col_idx == 1:
                        row_data.append(cell_text)  # Annex Ref. (Excel col 1)
                    elif col_idx == 2:
                        row_data.append(cell_text)  # Standard (Excel col 2)
                    elif col_idx == 3:
                        row_data.append(cell_text)  # State Ref. (Excel col 4)
                    elif col_idx == 10:
                        row_data.append(cell_text)  # Details (Excel col 5)
                    elif col_idx == 11:
                        row_data.append(cell_text)  # Remark (Excel col 6)

            # Determine text for checked checkboxes
            if len(checked_indices) == 0:
                checked_text = ""  # No checkboxes checked
            elif len(checked_indices) == 1:
                checked_text = checkbox_text_map.get(checked_indices[0], "")
            else:
                checked_text = "error-multi checkbox"

            # Insert checked text as the third column (Excel col 3)
            row_data.insert(2, checked_text)  # Inserts "Difference" at index 2
            table_data.append(row_data)
            root.update()  # Update GUI after each row

        # Define exactly 6 column headers for Excel
        headers = [
            "Annex Ref.",      # Word col 1
            "Standard",        # Word col 2
            "Difference",      # Checkboxes from Word cols 4-9
            "State Ref.",      # Word col 3
            "Details",         # Word col 10
            "Remark"           # Word col 11
        ]

        # Create DataFrame with exactly 6 columns
        df = pd.DataFrame(table_data, columns=headers)

        # Generate unique output filename
        base_name = "output.xlsx"
        output_excel_path = os.path.join(output_dir, base_name)
        counter = 1
        while os.path.exists(output_excel_path):
            output_excel_path = os.path.join(output_dir, f"output_{counter}.xlsx")
            counter += 1

        # Export to Excel initially
        df.to_excel(output_excel_path, index=False, engine='openpyxl')

        # Load the workbook to modify it
        wb = load_workbook(output_excel_path)
        ws = wb.active

        # Define the table range (A1 to F<rows+1> for 6 columns)
        num_rows = len(table_data) + 1  # +1 for header
        table_range = f"A1:F{num_rows}"

        # Create an Excel table
        tab = Table(displayName="FormDataTable", ref=table_range)
        style = TableStyleInfo(name="TableStyleMedium2", showFirstColumn=False,
                               showLastColumn=False, showRowStripes=True, showColumnStripes=False)
        tab.tableStyleInfo = style
        ws.add_table(tab)

        # Freeze the first row (headers)
        ws.freeze_panes = ws['A2']  # Freezes row 1

        # Save the modified workbook
        wb.save(output_excel_path)
        logging.info(f"Table data exported to {output_excel_path} (Processed {len(table_data)} rows) with table and frozen headers")
        root.update()
        return output_excel_path

    except Exception as e:
        logging.error(f"An error occurred in export_table_to_excel: {e}")
        return None

    finally:
        try:
            doc.Close()
            logging.info("Document closed")
            root.update()
        except:
            pass
        try:
            word.Quit()
            logging.info("Word application quit")
            root.update()
        except:
            pass

def fill_form_from_excel(excel_path, form_path, root):
    if not os.path.exists(excel_path):
        logging.error(f"Excel file not found: {excel_path}")
        return None
    if not os.path.exists(form_path):
        logging.error(f"Form file not found: {form_path}")
        return None

    # Read Excel data
    try:
        df = pd.read_excel(excel_path)
        if list(df.columns) != ["Annex Ref.", "Standard", "Difference", "State Ref.", "Details", "Remark"]:
            logging.error("Excel file does not match expected column structure")
            return None
        logging.info("Excel data loaded successfully")
        root.update()
    except Exception as e:
        logging.error(f"Failed to read Excel file: {e}")
        return None

    # Initialize Word application
    try:
        word = win32.Dispatch('Word.Application')
        word.Visible = False
        word.DisplayAlerts = False
        logging.info("Word application initialized")
        root.update()
    except Exception as e:
        logging.error(f"Failed to initialize Word: {e}")
        return None

    try:
        # Open the document
        logging.debug(f"Attempting to open document: {form_path}")
        doc = word.Documents.Open(os.path.abspath(form_path))
        logging.info(f"Opened document: {form_path}")
        root.update()

        # Define Word constants
        WD_NO_PROTECTION = -1
        wdAllowOnlyFormFields = 2  # Protection type for form fields only

        # Check protection type (for logging/info purposes, no unprotection)
        if doc.ProtectionType == WD_NO_PROTECTION:
            logging.info("Document is not protected")
        elif doc.ProtectionType == wdAllowOnlyFormFields:
            logging.info("Document is protected with form fields only - editing form fields directly")
        else:
            logging.error(f"Document has unsupported protection type: {doc.ProtectionType}. Must be unprotected or form-fields-only.")
            doc.Close()
            return None
        root.update()

        logging.debug("Checking table count")
        if doc.Tables.Count == 0:
            logging.error("No tables found in the document.")
            doc.Close()
            return None

        # Get the first table
        logging.debug("Accessing first table")
        table = doc.Tables(1)

        # Verify Word table has 11 columns
        logging.debug("Verifying column count")
        if table.Columns.Count != 11:
            logging.error(f"Expected 11 columns in Word table, found {table.Columns.Count}")
            doc.Close()
            return None

        # Check if number of rows matches
        excel_rows = len(df)
        word_rows = table.Rows.Count
        if excel_rows != word_rows:
            logging.error(f"Row count mismatch: Excel has {excel_rows} rows, Word table has {word_rows} rows")
            messagebox.showerror("Error", f"Row count mismatch: Excel has {excel_rows} rows, Word table has {word_rows} rows")
            doc.Close()
            return None
        logging.info(f"Row count matches: {excel_rows} rows in both Excel and Word table")
        root.update()

        # Checkbox mapping with first 3 letters (case-insensitive)
        checkbox_text_map = {
            'nod': 4,  # No Difference
            'mor': 5,  # More Exacting
            'dif': 6,  # Different in character
            'les': 7,  # Less protective or partially
            'sig': 8,  # Significant Difference
            'not': 9   # Not Applicable
        }

        # Iterate through rows
        max_rows = table.Rows.Count
        # max_rows = min(30, table.Rows.Count, len(df))  # Still respects the 30-row limit
        logging.info(f"Processing up to {max_rows} rows")
        root.update()
        for row_idx in range(1, max_rows + 1):
            row_data = df.iloc[row_idx - 1]  # 0-based index for pandas

            # Skip modifying columns 1 and 2 (as per original requirement)
            # Fill columns 3, 10, 11 (assuming these are text form fields)
            for col_idx, value in [(3, row_data["State Ref."]), (10, row_data["Details"]), (11, row_data["Remark"])]:
                try:
                    cell = table.Cell(row_idx, col_idx)
                    if cell.Range.FormFields.Count > 0:
                        # Use the first form field in the cell (assuming text field)
                        field = cell.Range.FormFields(1)
                        if field.Type == 70:  # wdFieldFormTextInput
                            field.Result = str(value) if pd.notna(value) else ""
                            logging.debug(f"Set text form field in Row {row_idx}, Col {col_idx}: {value}")
                        else:
                            logging.warning(f"Expected text form field, found type {field.Type} in Row {row_idx}, Col {col_idx}")
                    else:
                        logging.warning(f"No form field found in Row {row_idx}, Col {col_idx} - skipping text update")
                except Exception as e:
                    logging.error(f"Failed to set text in Row {row_idx}, Col {col_idx}: {e}")
            root.update()  # Update after text fields

            # Handle checkboxes (columns 4-9): Reset all to unchecked first
            for col_idx in range(4, 10):  # Columns 4-9
                try:
                    cell = table.Cell(row_idx, col_idx)
                    if cell.Range.FormFields.Count == 0:
                        logging.warning(f"No form fields in Row {row_idx}, Col {col_idx}")
                        continue
                    for field in cell.Range.FormFields:
                        if field.Type == 71:  # wdFieldFormCheckBox
                            field.CheckBox.Value = False  # Uncheck all boxes
                            logging.debug(f"Unchecked checkbox in Row {row_idx}, Col {col_idx}")
                        else:
                            logging.warning(f"Unexpected field type {field.Type} in Row {row_idx}, Col {col_idx}")
                except Exception as e:
                    logging.error(f"Error resetting checkbox in Row {row_idx}, Col {col_idx}: {e}")

            # Check the appropriate box based on Excel data using first 3 letters
            diff_value = row_data["Difference"]
            if pd.notna(diff_value) and diff_value != "error-multi checkbox":
                # Remove leading spaces and convert to lowercase, take first 3 letters
                diff_key = str(diff_value).lstrip().lower()[:3]
                if diff_key in checkbox_text_map:
                    col_idx = checkbox_text_map[diff_key]
                    try:
                        cell = table.Cell(row_idx, col_idx)
                        if cell.Range.FormFields.Count == 0:
                            logging.warning(f"No form fields to check in Row {row_idx}, Col {col_idx}")
                        for field in cell.Range.FormFields:
                            if field.Type == 71:  # wdFieldFormCheckBox
                                field.CheckBox.Value = True
                    except Exception as e:
                        logging.error(f"Error checking checkbox in Row {row_idx}, Col {col_idx}: {e}")
            root.update()

        # Save the modified document with a new name
        output_form_path = os.path.splitext(form_path)[0] + "_edited.docx"
        counter = 1
        while os.path.exists(output_form_path):
            output_form_path = os.path.splitext(form_path)[0] + f"_edited_{counter}.docx"
            counter += 1

        logging.debug(f"Saving document as: {output_form_path}")
        doc.SaveAs(os.path.abspath(output_form_path))
        logging.info(f"Form filled and saved as {output_form_path}")
        root.update()
        return output_form_path

    except Exception as e:
        logging.error(f"An error occurred in fill_form_from_excel: {e}")
        return None

    finally:
        try:
            doc.Close()
            logging.info("Document closed")
            root.update()
        except:
            pass
        try:
            word.Quit()
            logging.info("Word application quit")
            root.update()
        except:
            pass

def xml_to_excel(xml_path, output_dir, root):
    if not os.path.exists(xml_path):
        logging.error(f"XML file not found: {xml_path}")
        return None

    try:
        # Parse the XML file
        tree = ET.parse(xml_path)
        xml_root = tree.getroot()

        # Define namespaces
        namespaces = {'ns': 'urn:crystal-reports:schemas:report-detail'}

        # Extract relevant data
        table_data = []
        for details in xml_root.findall('.//ns:Details', namespaces):
            row_data = []
            annex_ref = details.find('.//ns:Field[@Name="AnnexReferenceNumber1"]/ns:Value', namespaces)
            standard = details.find('.//ns:Field[@Name="SARP11"]/ns:Value', namespaces)
            state_ref = details.find('.//ns:Field[@Name="StateReference1"]/ns:Value', namespaces)
            difference = details.find('.//ns:Field[@Name="StateDifferenceLevel1"]/ns:Value', namespaces)
            state_difference = details.find('.//ns:Field[@Name="StateDifference1"]/ns:Value', namespaces)
            state_comments = details.find('.//ns:Field[@Name="StateComments1"]/ns:Value', namespaces)

            # Check if elements are found and extract text
            annex_ref_text = annex_ref.text if annex_ref is not None else "Not Found"
            standard_text = standard.text if standard is not None else "Not Found"
            state_ref_text = state_ref.text if state_ref is not None else "Not Found"
            difference_text = difference.text if difference is not None else "Not Found"
            state_difference_text = state_difference.text if state_difference is not None else "Not Found"
            state_comments_text = state_comments.text if state_comments is not None else "Not Found"

            # Log the extracted data
            logging.debug(f"Annex Ref: {annex_ref_text}, Standard: {standard_text}, State Ref: {state_ref_text}, "
                          f"Difference: {difference_text}, Details: {state_difference_text}, Remark: {state_comments_text}")

            row_data.append(annex_ref_text)
            row_data.append(standard_text)
            row_data.append(difference_text)
            row_data.append(state_ref_text)
            row_data.append(state_difference_text)
            row_data.append(state_comments_text)

            table_data.append(row_data)

        if not table_data:
            logging.error("No data extracted from XML.")
            return None

        # Define column headers for Excel
        headers = [
            "Annex Ref.",
            "Standard",
            "Difference",
            "State Ref.",
            "Details",
            "Remark"
        ]

        # Create DataFrame
        df = pd.DataFrame(table_data, columns=headers)
        logging.debug(f"DataFrame created with data:\n{df}")

        # Generate unique output filename
        base_name = "output_from_xml.xlsx"
        output_excel_path = os.path.join(output_dir, base_name)
        counter = 1
        while os.path.exists(output_excel_path):
            output_excel_path = os.path.join(output_dir, f"output_from_xml_{counter}.xlsx")
            counter += 1

        # Export to Excel initially
        df.to_excel(output_excel_path, index=False, engine='openpyxl')

        # Load the workbook to modify it
        wb = load_workbook(output_excel_path)
        ws = wb.active

        # Define the table range (A1 to F<rows+1> for 6 columns)
        num_rows = len(table_data) + 1  # +1 for header
        table_range = f"A1:F{num_rows}"

        # Create an Excel table
        tab = Table(displayName="FormDataTable", ref=table_range)
        style = TableStyleInfo(name="TableStyleMedium2", showFirstColumn=False,
                               showLastColumn=False, showRowStripes=True, showColumnStripes=False)
        tab.tableStyleInfo = style
        ws.add_table(tab)

        # Freeze the first row (headers)
        ws.freeze_panes = ws['A2']  # Freezes row 1

        # Save the modified workbook
        wb.save(output_excel_path)
        logging.info(f"XML data exported to {output_excel_path} (Processed {len(table_data)} rows) with table and frozen headers")
        root.update()
        return output_excel_path

    except Exception as e:
        logging.error(f"An error occurred in xml_to_excel: {e}")
        return None

def gui():
    root = tk.Tk()
    root.title("EFOD Converter")
    root.geometry("800x400")  # Initial size

    # Frame for buttons
    button_frame = tk.Frame(root)
    button_frame.pack(pady=10, fill=tk.X)  # Fill horizontally

    # Log display area (dynamic width)
    log_text = scrolledtext.ScrolledText(root, height=20, state='disabled')
    log_text.pack(pady=10, fill=tk.BOTH, expand=True)  # Fill both directions and expand

    # Set up logging to the text widget
    setup_logging(log_text, root)

    def form_to_excel():
        form_path = filedialog.askopenfilename(title="Select Word Form", filetypes=[("Word files", "*.docx")])
        if form_path:
            output_dir = os.path.dirname(form_path)
            output_file = export_table_to_excel(form_path, output_dir, root)
            if output_file:
                messagebox.showinfo("Success", f"Conversion completed. Output saved as: {output_file}")
            else:
                messagebox.showerror("Error", "Conversion failed. Check logs for details.")

    def excel_to_form():
        excel_path = filedialog.askopenfilename(title="Select Excel File", filetypes=[("Excel files", "*.xlsx")])
        if excel_path:
            form_path = filedialog.askopenfilename(title="Select EFOD Form to Edit", filetypes=[("Word files", "*.docx")])
            if form_path:
                output_file = fill_form_from_excel(excel_path, form_path, root)
                if output_file:
                    messagebox.showinfo("Success", f"Form filled and saved as: {output_file}")
                else:
                    messagebox.showerror("Error", "Conversion failed. Check logs for details.")

    def xml_to_excel_conversion():
        xml_path = filedialog.askopenfilename(title="Select XML File of a country, Exported from SAP Crystal Reports", filetypes=[("XML files", "*.xml")])
        if xml_path:
            output_dir = os.path.dirname(xml_path)
            output_file = xml_to_excel(xml_path, output_dir, root)
            if output_file:
                messagebox.showinfo("Success", f"Conversion completed. Output saved as: {output_file}")
            else:
                messagebox.showerror("Error", "Conversion failed. Check logs for details.")

    btn_form_to_excel = tk.Button(button_frame, text="EFOD to Excel", command=form_to_excel, width=20)
    btn_form_to_excel.pack(side=tk.LEFT, padx=10)

    btn_excel_to_form = tk.Button(button_frame, text="Excel to EFOD", command=excel_to_form, width=20)
    btn_excel_to_form.pack(side=tk.LEFT, padx=10)

    btn_xml_to_excel = tk.Button(button_frame, text="SAP Crystal Reports to Excel", command=xml_to_excel_conversion, width=20)
    btn_xml_to_excel.pack(side=tk.LEFT, padx=10)

    root.mainloop()

if __name__ == "__main__":
    gui()
