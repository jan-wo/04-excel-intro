import xlwings as xw
from xlwings import constants
import pandas as pd


class ExcelFile():

    def __init__(self, visible=False):
        print("[ExcelFile]: object created")
        self._app = xw.App(visible=visible)

    def stop_app(self):
        print(f'[ExcelFile]: Closing application.')
        self._app.quit()

    def add_workbook(self, wb_path):
        print(f'[ExcelFile]: Adding workbook: <{wb_path}>')
        self._wb = self._app.books.open(wb_path)

        print(f'[ExcelFile]: Workbooks assigned to app: {self._app.books}')

    def save_close_wb(self, new_wb_path):
        print(f'[ExcelFile]: Saving application as <{new_wb_path}>')
        self._wb.save(new_wb_path)
        self._wb.close()

    def select_sheet(self, sheet_name):
        self._sh = self._wb.sheets(sheet_name)
        print(f'[ExcelFile]: Selected sheet <{self._sh}>')

    def data_from_range(self, address):
        data = self._sh.range(address).value
        print(f'[ExcelFile]: Data taken from <{address}>: <{data}>')

    def data_to_range(self, data, address):
        self._sh.range(address).value = data
        print(f'[ExcelFile]: Inserting <{data}> to <{address}>')


    def find_last_row(self, column_name):
        """
        This method find the last non-empty row in the column provided in
        parameter column_name and set it as _last_row attribute
        """
        # Find address of the most bottom cell in a given column.
        last_empty = column_name + str(self._sh.cells.last_cell.row)

        # Go from the bottom to the top, stop at first non-empty cell.
        last_row =  self._sh.range(last_empty).end('up').row
        print(f'[ExcelFile]: The last row found: <{last_row}>')
        return last_row


    def column_to_series(self, column_name, header=0):
        first_address = column_name + str(header + 1)
        last_address = column_name + str(self.find_last_row(column_name))
        column = self._sh.range(f'{first_address}:{last_address}')
        series_column = column.options(pd.Series, index=False).value
        return series_column

    def paint_range(self, range_address, rgb=None):
        """
        This method requires xlwings range address to be passed
        """
        # Paint in red if no rgb given
        if not rgb:
            rgb = (255, 0, 0)

        self._sh.range(range_address).color = rgb
