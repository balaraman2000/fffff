import csv
import xlwings as xw


class PastTheData:
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

            wb = xw.Book(self.target_file)
            ws = wb.sheets[1]

            for i, (key, value) in enumerate(data_dict.items(), start=3):
                ws.cells(i, 11).value = key
                ws.cells(i, 12).value = value

            wb.save()
            wb.close()
            print("Data replaced successfully.")
        except Exception as e:
            print(f"An error occurred: {e}")
