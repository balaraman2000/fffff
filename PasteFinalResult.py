import pandas as pd
import xlwings as xw
from datetime import datetime, timedelta


class PasteFinalResult:
    def __init__(self, source_file, target_file):
        self.source_file = source_file
        self.target_file = target_file
        self.copy_final_result()
        self.copy_final_result1()
        self.copy_final_result2()

    def copy_final_result(self):
        try:
            df = pd.read_excel(self.source_file, sheet_name='TTL miles driven(km)')
            column_data_list = df.iloc[:39, 10].dropna().tolist()
            wb = xw.Book(self.target_file)
            ws = wb.sheets[0]
            for i, value in enumerate(column_data_list, start=3):
                ws.cells(i, 3).value = value

            today = datetime.now()

            first_day_of_this_month = today.replace(day=1)

            last_day_of_last_month = first_day_of_this_month - timedelta(days=1)

            last_day_formatted = last_day_of_last_month.date().strftime("%d/%m/%Y") + " UTC"

            print("Last day of the previous month:", last_day_formatted)

            ws.range('C11').value = last_day_formatted

            wb.save()
            wb.close()


        except Exception as e:
            print("Error:", e)

    def copy_final_result1(self):
        try:
            df = pd.read_excel(self.source_file, sheet_name='Total CO2 Saved (kg)')
            column_data_list = df.iloc[:39, 8].dropna().tolist()
            wb = xw.Book(self.target_file)
            ws = wb.sheets[0]
            for i, value in enumerate(column_data_list, start=3):
                ws.cells(i, 5).value = value

            wb.save()
            wb.close()


        except Exception as e:
            print("Error:", e)

    def copy_final_result2(self):
        try:
            df = pd.read_excel(self.source_file, sheet_name='Num of vehicle')
            column_data_list = df.iloc[1:15, 7].dropna().tolist()

            wb = xw.Book(self.target_file)
            ws = wb.sheets[0]
            for i, value in enumerate(column_data_list, start=3):
                ws.cells(i, 6).value = value

            wb.save()
            wb.close()

            print("Data copied successfully.")
        except Exception as e:
            print("Error:", e)
