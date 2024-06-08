from openpyxl import load_workbook
from pyexcel.cookbook import merge_all_to_a_book
import glob
import os
from robot.api import logger

class CExcel(object):
    def get_number_of_cols_in_excel(self, filePath):
        numOfCols = 0
        file = load_workbook(filePath)
        sheet = file.active
        numOfCols = sheet.max_column
        return numOfCols
    def get_number_of_rows_in_excel(self, filePath):
        numOfRows = 0
        file = load_workbook(filePath)
        sheet = file.active
        numOfRows = sheet.max_row
        return numOfRows
    def convert_csv_to_xlsx(self, csvFilePath, xlsxFilePath):
        # logger.console(csvFilePath)
        # logger.console(xlsxFilePath)
        merge_all_to_a_book(glob.glob(csvFilePath), xlsxFilePath)
        os.remove(csvFilePath)

    def get_position_of_column(self, filePath, rowIndex, searchStr):
        posOfColumn = 0
        numOfCols = self.get_number_of_cols_in_excel(filePath)
        file = load_workbook(filePath)
        sheet = file.active
        # valueOfCell = sheet.cell(row=4, column=5).value
        # logger.console("valueOfCell: {0}".format(valueOfCell))
        for colIndex in range(1, numOfCols+1):
            valueOfCell = sheet.cell(row=rowIndex, column=colIndex).value
            if valueOfCell == searchStr:
                posOfColumn = colIndex
                break
        return posOfColumn





