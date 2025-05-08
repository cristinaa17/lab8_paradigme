import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.*;

import java.io.FileOutputStream;
import java.util.Map;
import java.util.TreeMap;

public class Problema2 {

    public static void main(String[] args) {
        XSSFWorkbook workbook = new XSSFWorkbook();
        XSSFSheet sheet = workbook.createSheet("Grades");

        XSSFCellStyle headerStyle = workbook.createCellStyle();
        XSSFFont headerFont = workbook.createFont();
        headerFont.setBold(true);
        headerFont.setFontName("Arial");
        headerFont.setFontHeightInPoints((short) 12);
        headerStyle.setFont(headerFont);
        headerStyle.setFillForegroundColor(IndexedColors.LIGHT_GREEN.getIndex());
        headerStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

        XSSFCellStyle yellowStyle = workbook.createCellStyle();
        yellowStyle.setFillForegroundColor(IndexedColors.LIGHT_YELLOW.getIndex());
        yellowStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

        String[] header = {"Name", "Surname", "Grade 1", "Grade 2", "Grade 3", "Grade 4", "Max", "Average", "Median"};
        Row headerRow = sheet.createRow(0);
        for (int i = 0; i < header.length; i++) {
            Cell cell = headerRow.createCell(i);
            cell.setCellValue(header[i]);
            cell.setCellStyle(headerStyle);
        }

        Map<String, Object[]> data = new TreeMap<>();
        data.put("2", new Object[]{"Amit", "Shukla", 9, 8, 7, 5});
        data.put("3", new Object[]{"Lokesh", "Gupta", 8, 9, 6, 7});
        data.put("4", new Object[]{"John", "Adwards", 8, 8, 7, 6});
        data.put("5", new Object[]{"Brian", "Schultz", 7, 6, 8, 9});

        int rowNum = 1;
        for (String key : data.keySet()) {
            Row row = sheet.createRow(rowNum);
            Object[] objArr = data.get(key);
            int colNum = 0;
            for (Object obj : objArr) {
                Cell cell = row.createCell(colNum++);
                if (obj instanceof String)
                    cell.setCellValue((String) obj);
                else if (obj instanceof Integer)
                    cell.setCellValue((Integer) obj);
            }

            String formulaMax = "MAX(C" + (rowNum + 1) + ":F" + (rowNum + 1) + ")";
            Cell cellMax = row.createCell(6);
            cellMax.setCellFormula(formulaMax);
            cellMax.setCellStyle(yellowStyle);

            String formulaAvg = "AVERAGE(C" + (rowNum + 1) + ":F" + (rowNum + 1) + ")";
            Cell cellAvg = row.createCell(7);
            cellAvg.setCellFormula(formulaAvg);
            cellAvg.setCellStyle(yellowStyle);

            String formulaMed= "MEDIAN(C" + (rowNum + 1) + ":F" + (rowNum + 1) + ")";
            Cell cellMed = row.createCell(8);
            cellMed.setCellFormula(formulaMed);
            cellMed.setCellStyle(yellowStyle);

            rowNum++;
        }

        Row sumRow = sheet.createRow(rowNum);
        Cell totalLabelCell = sumRow.createCell(0);
        totalLabelCell.setCellValue("Total");
        totalLabelCell.setCellStyle(headerStyle);


        sumRow.createCell(1).setCellValue("");

        char colLetter = 'C';
        for (int col = 2; col <= 8; col++) {
            String formula = "SUM(" + colLetter + "2:" + colLetter + (rowNum) + ")";
            Cell sumCell = sumRow.createCell(col);
            sumCell.setCellFormula(formula);
            sumCell.setCellStyle(yellowStyle);
            colLetter++;
        }


        for (int i = 0; i < header.length; i++) {
            sheet.autoSizeColumn(i);
        }

        try (FileOutputStream out = new FileOutputStream("output8.xlsx")) {
            workbook.write(out);
            workbook.close();
            System.out.println("Fisierul output8.xlsx a fost generat cu succes.");
        } catch (Exception e) {
            e.printStackTrace();
        }
    }
}
