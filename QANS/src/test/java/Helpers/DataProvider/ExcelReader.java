package Helpers.DataProvider;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.ArrayList;
import java.util.List;

public class ExcelReader {

    public ExcelReader(){

    }

    public List<Object[]> readDataFromExcel(String filePath, int sheetIndex, int startRow, int headerRowIndex){
        List<Object[]> rowsArr = new ArrayList<Object[]>();
        try
        {
            rowsArr.clear();
            File file = new File(filePath);
            FileInputStream fileInputStream = new FileInputStream(file);
            XSSFWorkbook wb = new XSSFWorkbook(fileInputStream);
            XSSFSheet sheet = wb.getSheetAt(sheetIndex);
            int numOfRows = sheet.getLastRowNum();
            int numOfCols = sheet.getRow(headerRowIndex).getPhysicalNumberOfCells();
            for (int rowIndex = startRow; rowIndex <= numOfRows; rowIndex++){
                Row row = sheet.getRow(rowIndex);
                Object[] colsArr = new Object[numOfCols];
                for (int colIndex = 0; colIndex < numOfCols; colIndex++){
                    Cell cell = row.getCell(colIndex);
                    if(cell == null){
                        cell = row.getCell(colIndex, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
                    }
                    switch (cell.getCellType()){
                        case STRING:
                            colsArr[colIndex] = cell.getStringCellValue();
                            break;
                        case NUMERIC:
                            colsArr[colIndex] = cell.getNumericCellValue();
                            break;
                        case BLANK:
                            colsArr[colIndex] = "";
                            break;
                        case BOOLEAN:
                            colsArr[colIndex] = cell.getBooleanCellValue();
                            break;
                        case FORMULA:
                            colsArr[colIndex] = cell.getCellFormula().toString();
                            break;
                        default:
                            throw new IllegalArgumentException(String.format("The format of cell at row %d and column %d is not supported", rowIndex, colIndex));
                    }
                }
                rowsArr.add(colsArr);
            }
            fileInputStream.close();
        }catch (Exception e){
            System.out.println(String.format("Error: %s", e.getMessage()));
        }
        return rowsArr;
    }

    public void writeDataToExcel(String filePath, int sheetIndex, int rowIndex, int colIndex, Object cellValue){
        try{
            File file = new File(filePath);
            FileInputStream fileInputStream = new FileInputStream(file);
            XSSFWorkbook wb = new XSSFWorkbook(fileInputStream);
            XSSFSheet sheet = wb.getSheetAt(sheetIndex);
            Row row = null;
            if (sheet.getRow(rowIndex) != null){
                row = sheet.getRow(rowIndex);
            }else {
                row = sheet.createRow(rowIndex);
            }
            Cell cell = row.createCell(colIndex);
            if (cellValue instanceof Double){
                cell.setCellValue(Double.valueOf(cellValue.toString()));
            }else if (cellValue instanceof String){
                cell.setCellValue(cellValue.toString());
            }
            FileOutputStream fout = new FileOutputStream(file);
            wb.write(fout);
            fout.close();

        }catch (Exception e){
            System.out.println(String.format("Error: %s", e.getMessage()));
        }
    }


}
