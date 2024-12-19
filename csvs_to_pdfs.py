import os
import re
import win32com.client

def save_excel_sheets_as_pdfs(file_path):
    if not os.path.exists(file_path):
        print(f"Error: File '{file_path}' does not exist.")
        return

    try:
        excel_app = win32com.client.Dispatch("Excel.Application")
        excel_app.Visible = False
        excel_app.DisplayAlerts = False  # Disable Excel alerts

        workbook = excel_app.Workbooks.Open(file_path)

        output_dir = os.path.abspath("output_pdfs")
        os.makedirs(output_dir, exist_ok=True)

        for sheet in workbook.Sheets:
            original_sheet_name = sheet.Name
            # Sanitize sheet name for file saving
            sheet_name = original_sheet_name.strip()
            # Remove invalid filename characters
            sheet_name = re.sub(r'[\\/*?:<>|"]', "", sheet_name)

            pdf_filename = f"REP_{sheet_name}.pdf"
            pdf_filepath = os.path.join(output_dir, pdf_filename)

            # Adjust page setup if necessary
            sheet.PageSetup.Zoom = False
            sheet.PageSetup.FitToPagesTall = 1
            sheet.PageSetup.FitToPagesWide = 1

            print(f"Accessing sheet: {original_sheet_name}")

            try:
                sheet.ExportAsFixedFormat(
                    Type=0,  # xlTypePDF
                    Filename=pdf_filepath,
                    Quality=0,  # xlQualityStandard
                    IncludeDocProperties=True,
                    IgnorePrintAreas=True,
                    OpenAfterPublish=False
                )
                print(f"Saved: {pdf_filepath}")
            except Exception as e:
                print(f"Failed to process sheet '{original_sheet_name}': {e}")

        workbook.Close(SaveChanges=False)
        excel_app.Quit()

        print(f"All sheets have been processed. Check '{output_dir}' for PDFs.")

    except Exception as e:
        print(f"An error occurred: {e}")

# Example usage
if __name__ == "__main__":
    # Replace 'your_excel_file.xlsx' with the actual path to your Excel file
    excel_file_path = r"File NAME .xslx HERE"
    save_excel_sheets_as_pdfs(excel_file_path)
