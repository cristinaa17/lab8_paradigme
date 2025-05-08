import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;


public class Main {
    public static void main(String[] args) {
        XSSFWorkbook workbook = null;
        try {
            FileInputStream file = new FileInputStream(new
                    File("laborator8_input.xlsx"));
//Create Workbook instance holding reference to .xlsx file
            workbook = new XSSFWorkbook(file);
//Get first/desired sheet from the workbook
            XSSFSheet sheet = workbook.getSheetAt(0);
//Iterate through each rows one by one
            for (Row row : sheet) {
//For each row, iterate through all the columns
                for (Cell cell : row) {
//Check the cell type and format accordingly
                    switch (cell.getCellType()) {
                        case CellType.NUMERIC:
                            System.out.print(cell.getNumericCellValue() + "\t");
                            break;
                        case CellType.STRING:
                            System.out.print(cell.getStringCellValue() + "\t");
                            break;
                    }
                }
                System.out.println();
            }
            file.close();
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
    }
}
