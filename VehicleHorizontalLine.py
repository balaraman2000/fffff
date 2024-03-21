import xlwings as xw
import win32com.client as win32
from datetime import datetime, timedelta


class VehicleHorizontalLine:
    def __init__(self,source_file,target_file):
        self.source_file = source_file
        self.target_file = target_file
        self.copy_and_paste_horizontalLine_excel()
        self.find_last_index()

    def copy_and_paste_horizontalLine_excel(self):
        try:
            wb_source = xw.Book(self.source_file)
            wb_target = xw.Book(self.target_file)

            source_sheet = wb_source.sheets['Sheet1']
            target_sheet = wb_target.sheets[4]

            last_row_index = target_sheet.range("C" + str(target_sheet.cells.last_cell.row)).end('up').row

            source_range = source_sheet.range('F3:F8')
            target_range = target_sheet.range('C' + str(last_row_index + 1))


            for i, cell in enumerate(source_range):

                target_cell = target_range.offset(row_offset=0, column_offset=i)
                target_cell.value = cell.value
                target_cell.number_format = cell.number_format


            target_range.api.Resize(1, 6).Interior.Color = 0xFF99CC  # Pink color

            target_sheet.range('C' + str(last_row_index + 1) + ':H' + str(last_row_index + 1)).color = (
                255, 153, 204, 255)

            wb_target.save()

        except FileNotFoundError:
            print("File not found. Check the file paths.")
        except xw.exceptions.RangeError:
            print("Error accessing Excel range.")
        except Exception as e:
            print("An error occurred:", e)
        finally:
            wb_source.close()
            wb_target.close()

    def find_last_index(self):
        excel = None
        try:
            excel = win32.Dispatch("Excel.Application")
            excel.Visible = False
            workbook = excel.Workbooks.Open(self.target_file)
            sheet = workbook.Sheets("Num of vehicle")

            last_index = sheet.Cells.SpecialCells(11).Row
            before_last_index = last_index - 1

            C = "C" + str(last_index)
            C1 = "C" + str(before_last_index)
            formulaC = f'= {C} - {C1}'
            sheet.Range(f'K{last_index}').Formula = formulaC

            D = "D" + str(last_index)
            D1 = "D" + str(before_last_index)
            formulaD = f'= {D} - {D1}'
            sheet.Range(f'L{last_index}').Formula = formulaD

            E = "E" + str(last_index)
            E1 = "E" + str(before_last_index)
            formulaE = f'= {E} - {E1}'
            sheet.Range(f'M{last_index}').Formula = formulaE

            F = "F" + str(last_index)
            F1 = "F" + str(before_last_index)
            formulaF = f'= {F} - {F1}'
            sheet.Range(f'N{last_index}').Formula = formulaF

            G = "G" + str(last_index)
            G1 = "G" + str(before_last_index)
            formulaG = f'= {G} - {G1}'
            sheet.Range(f'O{last_index}').Formula = formulaG

            H = "H" + str(last_index)
            H1 = "H" + str(before_last_index)
            formulaH = f'= {H} - {H1}'
            sheet.Range(f'P{last_index}').Formula = formulaH

            border_range = sheet.Range(f'K{last_index}' + ":" + f'P{last_index}')
            border = border_range.Borders
            border.LineStyle = win32.constants.xlContinuous  # Continuous line style
            border.Weight = win32.constants.xlThin

            border_range = sheet.Range(f'C{last_index}' + ":" + f'H{last_index}')
            border = border_range.Borders
            border.LineStyle = win32.constants.xlContinuous  # Continuous line style
            border.Weight = win32.constants.xlThin

            today = datetime.now()

            first_day_of_this_month = today.replace(day=1)

            last_day_of_last_month = first_day_of_this_month - timedelta(days=1)

            last_day_formatted = last_day_of_last_month.date().strftime("%Y/%m/%d")
            sheet.Range(f'B{last_index}').Formula = last_day_formatted
            sheet.Range(f'j{last_index}').Formula = last_day_formatted

            workbook.Save()


        except Exception as e:
            print("Error:", e)

        finally:
            if excel is not None:
                excel.Quit()



