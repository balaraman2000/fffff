import win32com.client

def delete_columns():
    # Path to your Excel file
    file_path = r"E:\\AMO 20\\集計.xlsx"

    # Create an instance of Excel application
    excel = win32com.client.Dispatch("Excel.Application")

    # Open the workbook
    workbook = excel.Workbooks.Open(file_path)

    try:
        # Get the sheet
        sheet = workbook.Sheets("TTL miles driven(km)")  # Replace "Sheet2" with your sheet name

        # Define the range to delete (A2:B38)
        delete_range = sheet.Range("A2:B38")

        # Delete the range
        delete_range.Delete()

        # Save changes
        workbook.Save()

    except Exception as e:
        print("An error occurred:", e)

    finally:
        # Close the workbook and quit Excel application
        workbook.Close(SaveChanges=True)
        excel.Quit()

# Call the function to delete the columns
delete_columns()
