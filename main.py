import os
import win32com.client as win32
import pandas as pd
import re
from openpyxl import load_workbook
from openpyxl.worksheet.table import Table, TableStyleInfo


def export_table_to_excel(file_path, output_excel_path):
    if not os.path.exists(file_path):
        print(f"File not found: {file_path}")
        return

    # Initialize Word application
    word = win32.Dispatch('Word.Application')
    word.Visible = False
    word.DisplayAlerts = False

    try:
        # Open the document
        doc = word.Documents.Open(os.path.abspath(file_path))

        # Define Word constants
        WD_NO_PROTECTION = -1

        # Handle protection
        protection_password = None  # Set to your password if known
        if doc.ProtectionType != WD_NO_PROTECTION:
            print("Document is protected. Attempting to unprotect...")
            try:
                if protection_password:
                    doc.Unprotect(protection_password)
                else:
                    doc.Unprotect()
                print("Document unprotected successfully.")
            except Exception as e:
                print(f"Failed to unprotect document: {e}")
                doc.Close()
                return

        if doc.Tables.Count == 0:
            print("No tables found in the document.")
            doc.Close()
            return

        # Get the first table
        table = doc.Tables(1)

        # Verify Word table has 11 columns
        if table.Columns.Count != 11:
            print(f"Expected 11 columns in Word table, found {table.Columns.Count}")
            doc.Close()
            return

        # Prepare data structure
        table_data = []
        checkbox_columns = range(4, 9)  # Columns 4-8 (1-based: 4, 5, 6, 7, 8)

        # Sample text mapping for checkbox indices
        checkbox_text_map = {
            '4': 'No Difference',
            '5': 'More Exacting',
            '6': 'Different in character',
            '7': 'Less protective or partially',
            '8': 'Significant Difference'
        }

        # Iterate through rows (1-based indexing)
        max_rows = table.Rows.Count
        max_rows = min(30, table.Rows.Count)  # TEMPORARY: Limit to first 30 rows; comment out for all rows
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
                    if ord(char) < 32 and char != '\n':  # Stop at control chars like \r, \x0b
                        break
                    visible_text += char
                visible_text = visible_text.strip()
                # If in first column, extract numeric part from visible text
                if col_idx == 1:
                    numeric_match = re.match(r'^\d+(?:\.\d+)?(?![.\d])', visible_text)
                    if numeric_match:
                        cell_text = numeric_match.group(0)
                        print(
                            f"Row {row_idx}, Col {col_idx} raw: {repr(raw_text)}, visible: {repr(visible_text)}, matched: {cell_text}, remainder: {repr(raw_text[len(visible_text):].strip())}")
                    else:
                        cell_text = visible_text  # Fallback to visible text if no match
                        print(
                            f"Row {row_idx}, Col {col_idx} raw: {repr(raw_text)}, visible: {repr(visible_text)}, cleaned: {cell_text} (no numeric match)")
                else:
                    cell_text = visible_text  # Use visible text for other columns

                # Handle checkbox columns (4-8)
                if col_idx in checkbox_columns:
                    for field in cell.Range.FormFields:
                        if field.Type == 71:  # wdFieldFormCheckBox
                            if field.CheckBox.Value:
                                checked_indices.append(str(col_idx))
                            # Debug: Print raw cell content if problematic
                            if any(ord(c) < 32 and c != '\n' for c in raw_text):
                                print(f"Row {row_idx}, Col {col_idx} raw content: {repr(raw_text)}")
                # Add specific columns to row_data
                elif col_idx in [1, 2, 9, 10, 11]:  # Only these go to Excel
                    if col_idx == 1:
                        row_data.append(cell_text)  # Item Number
                    elif col_idx == 2:
                        row_data.append(cell_text)  # Description
                    elif col_idx >= 9:
                        row_data.append(cell_text)  # Category, Status, Notes

            # Determine text for checked checkboxes
            if len(checked_indices) == 0:
                checked_text = ""  # No checkboxes checked
            elif len(checked_indices) == 1:
                checked_text = checkbox_text_map.get(checked_indices[0], "")
            else:
                checked_text = "error-multi checkbox"

            # Insert checked text as the third column
            row_data.insert(2, checked_text)  # Puts "Checked Checkboxes" at index 2 (3rd column)

            table_data.append(row_data)

        # Define exactly 6 column headers for Excel
        headers = [
            "Annex Ref.",  # Word col 1
            "Standard",  # Word col 2
            "Difference",  # Generated from Word cols 4-8
            "State Ref.",  # Word col 9
            "Details",  # Word col 10
            "Remark"  # Word col 11
        ]

        # Create DataFrame with exactly 6 columns
        df = pd.DataFrame(table_data, columns=headers)

        # Export to Excel initially
        df.to_excel(output_excel_path, index=False, engine='openpyxl')

        # Load the workbook to modify it
        wb = load_workbook(output_excel_path)
        ws = wb.active

        # Define the table range (A1 to F31 for 30 rows + header)
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
        print(
            f"Table data exported to {output_excel_path} (Processed {len(table_data)} rows) with table and frozen headers")

    except Exception as e:
        print(f"An error occurred: {str(e)}")

    finally:
        try:
            doc.Close()
        except:
            pass
        try:
            word.Quit()
        except:
            pass


def main():
    file_path = os.path.join('form', 'form.docx')
    output_excel_path = os.path.join('form', 'form_data.xlsx')
    export_table_to_excel(file_path, output_excel_path)


if __name__ == "__main__":
    main()