package Helpers.Manager;

import Helpers.DataProvider.ExcelReader;
import Helpers.DataProvider.PropertyFileReader;

public class FileReaderManager {
    private static FileReaderManager fileReaderManager = new FileReaderManager();
    private PropertyFileReader propertyFileReader;
    private ExcelReader excelReader;
    private FileReaderManager(){

    }
    public static FileReaderManager getInstance(){
        return fileReaderManager;
    }
    public PropertyFileReader getPropertyFileReader(String filePath){
        return (propertyFileReader == null) ? propertyFileReader = new PropertyFileReader(filePath) : propertyFileReader;
    }
    public ExcelReader getExcelReader(){
        return (excelReader == null) ? excelReader = new ExcelReader() : excelReader;
    }
}
