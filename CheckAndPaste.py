import csv
import xlwings as xw
import pandas as pd

class CheckAndPaste:
    def __init__(self, source_file, target_file):
        self.source_file = source_file
        self.target_file = target_file
        self.delete_data_basic_on_D()

    def delete_data_basic_on_D(self):
        try:
            data_dict = {}
            with open(self.source_file, 'r') as file:
                reader = csv.reader(file)
                for row in reader:
                    if len(row) >= 2:
                        key = row[0]
                        try:
                            value = float(row[1])
                        except ValueError:
                            value = row[1]
                        data_dict[key] = value

        except FileNotFoundError:
            print(f"Error: File {self.source_file} not found.")

        try:
            df = pd.read_excel(self.target_file, sheet_name='TTL miles driven(km)')
            column_data_list = df.iloc[:37, 4].tolist()

        except Exception as e:
            print(f"Error occurred while reading Excel file: {e}")

        filtered_dict = {key: value for key, value in data_dict.items() if key in column_data_list}

        try:
            wb = xw.Book(self.target_file)
            sheet = wb.sheets['TTL miles driven(km)']
            for i, (key, value) in enumerate(filtered_dict.items(), start=2):
                sheet[f'A{i}'].value = key
                sheet[f'B{i}'].value = value
            wb.save()
            wb.close()
            print("Data appended successfully.")
        except FileNotFoundError:
            print(f"Error: Excel file {self.target_file} not found.")
        except Exception as e:
            print(f"Error occurred while opening Excel file: {e}")


