import os
import win32com.client as win32


def check_all_checkboxes_com(file_path):
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
        WD_ALLOW_ONLY_FORM_FIELDS = 2
        WD_FORMAT_XML_DOCUMENT = 12  # .docx format

        # Check if the document is protected
        was_protected = False
        protection_password = None  # Set to your password if known, e.g., "your_password"

        if doc.ProtectionType != WD_NO_PROTECTION:
            was_protected = True
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

        # Iterate through form fields and check only checkboxes
        for field in doc.FormFields:
            if field.Type == 71:  # wdFieldFormCheckBox = 71
                checkboxes_found = True
                field.CheckBox.Value = True
                print("Checked a legacy checkbox.")

        # Reapply protection if it was originally protected
        if was_protected:
            try:
                if protection_password:
                    doc.Protect(WD_ALLOW_ONLY_FORM_FIELDS, True, protection_password)
                else:
                    doc.Protect(WD_ALLOW_ONLY_FORM_FIELDS, True)
                print("Reapplied form protection.")
            except Exception as e:
                print(f"Failed to reapply protection: {e}")

        # Save the modified document
        output_path = os.path.join('form', 'modified_form.docx')
        doc.SaveAs(os.path.abspath(output_path), FileFormat=WD_FORMAT_XML_DOCUMENT)

        if checkboxes_found:
            print(f"Checkboxes checked. Modified file saved as: {output_path}")
        else:
            print("No form field checkboxes found in the document.")

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
    check_all_checkboxes_com(file_path)


if __name__ == "__main__":
    main()