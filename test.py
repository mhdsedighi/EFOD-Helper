import os
import win32com.client as win32


def analyze_checkboxes_in_table(file_path):
    if not os.path.exists(file_path):
        print(f"File not found: {file_path}")
        return

    # Initialize Word application
    word = win32.Dispatch('Word.Application')
    word.Visible = False  # Run in background
    word.DisplayAlerts = False  # Suppress alerts

    try:
        # Open the document
        doc = word.Documents.Open(os.path.abspath(file_path))

        # Define Word constants manually
        WD_NO_PROTECTION = -1

        # Check if the document is protected
        protection_password = None  # Set to your password if known, e.g., "your_password"

        if doc.ProtectionType != WD_NO_PROTECTION:
            print("Document is protected. Attempting to unprotect...")
            try:
                if protection_password:
                    doc.Unprotect(protection_password)
                else:
                    doc.Unprotect()  # Try without password
                print("Document unprotected successfully.")
            except Exception as e:
                print(f"Failed to unprotect document: {e}. Please provide password if applicable.")
                doc.Close()
                return

        checkboxes_found = False

        # Assume the document is one big table; get the first table
        if doc.Tables.Count == 0:
            print("No tables found in the document.")
            doc.Close()
            return

        table = doc.Tables(1)  # 1-based index in COM

        # Iterate through each row and column in the table
        for row_idx in range(1, table.Rows.Count + 1):
            for col_idx in range(1, table.Columns.Count + 1):
                cell = table.Cell(row_idx, col_idx)
                # Check for form fields in the cell's range
                for field in cell.Range.FormFields:
                    if field.Type == 71:  # wdFieldFormCheckBox = 71
                        checkboxes_found = True
                        # Check if the checkbox is already checked
                        checked_status = "checked" if field.CheckBox.Value else "unchecked"
                        print(f"Checkbox found at Row {row_idx}, Column {col_idx} - Status: {checked_status}")
                        # Do not modify: field.CheckBox.Value = True

        if not checkboxes_found:
            print("No checkboxes found in the table.")

        # No need to save since we're not modifying anything

    except Exception as e:
        print(f"An error occurred: {str(e)}")

    finally:
        # Clean up
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
    analyze_checkboxes_in_table(file_path)


if __name__ == "__main__":
    main()