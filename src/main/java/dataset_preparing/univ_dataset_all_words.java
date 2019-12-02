package dataset_preparing;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.*;

public class univ_dataset_all_words {
    public static final String FILE_PATH = "dataset_name_mtr_dedup.xlsx";

    public static void main(String args[]) {
        try {
            saveWorkbook(prepareExcel(loadExcel(FILE_PATH)));
        } catch (Exception e) {
            System.out.println(e.getMessage());
        }
    }

    private static TreeSet<String> loadExcel(String filename) throws IOException, InvalidFormatException {

        TreeSet<String> result = new TreeSet<>();

        Workbook workbook = WorkbookFactory.create(new File(filename));

        Sheet sheet = workbook.getSheetAt(0);

        DataFormatter dataFormatter = new DataFormatter();

        Iterator<Row> rowIterator = sheet.rowIterator();
        int i = 1;
        while (rowIterator.hasNext()) {
            Row row = rowIterator.next();

            Cell cell = row.getCell(0);
            String cellValue = dataFormatter.formatCellValue(cell);
            cellValue = cellValue.toLowerCase();
            cellValue = cellValue.replaceAll("[0-9A-Za-z;,.#()/*+-]", "");
            cellValue.replace("№", "");
            cellValue = cellValue.replaceAll(" {2,}", " ");
            String[] words;
            words = cellValue.split("\\s");
            for (String word: words) {
                if (word.length()>2) {
                    result.add(word);
                }
            }
            System.out.println("Record #" + i + " processed");
            i++;
        }

        return result;
    }

    private static Workbook prepareExcel(TreeSet<String> data) {
                XSSFWorkbook workbook = new XSSFWorkbook();

        XSSFSheet sheet = workbook.createSheet("univ_dataset_words");



        Iterator<String> data1Ie = data.iterator();
        Row row;
        Cell cell;
        int rownum = 0;

        while(data1Ie.hasNext()) {
            String word = data1Ie.next();
            row = sheet.createRow(rownum);
            cell = row.createCell(0);
            cell.setCellValue(word);
            rownum++;
            System.out.println("Word #"+ (rownum) + " processed.");
        }

        return workbook;



    }

    public static void saveWorkbook(Workbook workbook) {
        try {
            FileOutputStream out = new FileOutputStream(new File("univ_dataset_all_words.xlsx"));
            workbook.write(out);
            out.close();
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

}
