package dataset_preparing;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.HashSet;
import java.util.Iterator;

public class dataset_name_mtr_dedup {

    public static final String FILE_PATH = "data_for_univ_name_mtr.xlsx";




    public static void main(String args[]) {
        try {
            HashSet<String> excelData = loadExcel(FILE_PATH);
            saveToExcel(excelData);
        } catch (Exception e) {
            System.out.println(e.getMessage());
        }

    }

    private static HashSet<String> loadExcel(String filename) throws IOException, InvalidFormatException {

        HashSet<String> result = new HashSet<>();

        Workbook workbook = WorkbookFactory.create(new File(filename));

        Sheet sheet = workbook.getSheetAt(0);

        DataFormatter dataFormatter = new DataFormatter();

        Iterator<Row> rowIterator = sheet.rowIterator();

        while (rowIterator.hasNext()) {
            Row row = rowIterator.next();
            Cell cell = row.getCell(0);
            String cellValue = dataFormatter.formatCellValue(cell);
            result.add(cellValue);
        }

        return result;
    }

    private static void saveToExcel(HashSet<String> data) {
        XSSFWorkbook workbook = new XSSFWorkbook();

        XSSFSheet sheet = workbook.createSheet("name_mtr");

        Row firstRow = sheet.createRow(0);
        Cell cell = firstRow.createCell(0);
        cell.setCellValue("name_mtr");

        Iterator<String> rowIterator = data.iterator();
        int rownum = 1;
        while (rowIterator.hasNext()) {
            Row currentRow = sheet.createRow(rownum);
            String currentString = rowIterator.next();
            Cell currentCell = currentRow.createCell(0);
            currentCell.setCellValue(currentString);
            rownum++;
        }

        try {
            FileOutputStream out = new FileOutputStream(new File("dataset_name_mtr_dedup.xlsx"));
            workbook.write(out);
            out.close();
        } catch (Exception e) {
            e.printStackTrace();
        }
    }


}
