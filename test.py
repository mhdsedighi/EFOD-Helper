import os
import win32com.client as win32
import pandas as pd


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

        # Prepare data structure
        table_data = []
        checkbox_columns = range(4, 9)  # Columns 4 to 8 (1-based)

        # Sample text mapping for checkbox indices
        checkbox_text_map = {
            '4': 'No Difference',
            '5': 'More Exacting',
            '6': 'Different in character',
            '7': 'Less protective or partially',
            '8': 'Significant Difference'
        }

        # Iterate through rows (1-based indexing)
        max_rows = min(30, table.Rows.Count)  # TEMPORARY: Limit to first 30 rows; comment out for all rows
        for row_idx in range(1, max_rows + 1):  # Process up to max_rows
            row_data = []
            checked_indices = []

            # Iterate through columns
            for col_idx in range(1, table.Columns.Count + 1):
                cell = table.Cell(row_idx, col_idx)
                # Clean cell text: remove control chars and strip
                cell_text = ''.join(c for c in cell.Range.Text if ord(c) >= 32 or c == '\n').strip()

                # Handle checkbox columns (4-8)
                if col_idx in checkbox_columns:
                    for field in cell.Range.FormFields:
                        if field.Type == 71:  # wdFieldFormCheckBox
                            if field.CheckBox.Value:
                                checked_indices.append(str(col_idx))
                            # Debug: Print raw cell content if problematic
                            if any(ord(c) < 32 and c != '\n' for c in cell.Range.Text):
                                print(f"Row {row_idx}, Col {col_idx} raw content: {repr(cell.Range.Text)}")
                # Add non-checkbox column data
                elif col_idx < 4 or col_idx > 8:
                    if col_idx <= 2:
                        row_data.append(cell_text)  # Columns 1 and 2 first
                    elif col_idx > 8:
                        row_data.append(cell_text)  # Columns 9+ later

            # Determine text for checked checkboxes
            if len(checked_indices) == 0:
                checked_text = ""  # No checkboxes checked
            elif len(checked_indices) == 1:
                checked_text = checkbox_text_map.get(checked_indices[0], "")  # Single checkbox
            else:
                checked_text = "Error-multi checkbox"  # Multiple checkboxes

            # Insert checked text as the third column
            row_data.insert(2, checked_text)

            table_data.append(row_data)

        # Create column headers with "Checked Checkboxes" as third column
        headers = [f"Column {i}" for i in range(1, 3)] + ["Checked Checkboxes"] + \
                  [f"Column {i}" for i in range(3, 4)] + \
                  [f"Column {i}" for i in range(9, table.Columns.Count + 1)]

        # Create DataFrame
        df = pd.DataFrame(table_data, columns=headers[:len(table_data[0])])

        # Export to Excel
        df.to_excel(output_excel_path, index=False, engine='openpyxl')
        print(f"Table data exported to {output_excel_path} (Processed {len(table_data)} rows)")

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