from openpyxl import load_workbook
from pyexcel.cookbook import merge_all_to_a_book
import glob
import os
from robot.api import logger
from selenium import webdriver

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
        merge_all_to_a_book(glob.glob(csvFilePath), xlsxFilePath)
        os.remove(csvFilePath)
    def create_profile(self, path):
        logger.console(path)
        fp = webdriver.FirefoxProfile()
        fp.set_preference("browser.download.dir", path)
        fp.update_preferences()
        logger.console("Created profile")





