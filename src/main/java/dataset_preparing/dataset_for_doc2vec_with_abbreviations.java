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
import java.util.Objects;

public class dataset_for_doc2vec_with_abbreviations {

    public static final String FILE_PATH = "data_for_univ_name_mtr.xlsx";
    public static final String FILE_PATH2 = "abbreviations.xlsx";

    public static void main(String args[]) {
        try {

            saveWorkbook(prepareExcelWithAbbr(loadExcel(FILE_PATH)), "doc2vec_name_mtr_abbr.xlsx");

        } catch (Exception e) {
            System.out.println(e.getMessage());
        }
    }

    private static Workbook prepareExcelWithAbbr(HashSet<String> data) {
        HashSet<String> data2 = new HashSet<>(data);
        XSSFWorkbook workbook = new XSSFWorkbook();

        XSSFSheet sheet = workbook.createSheet("doc2vec_name_mtr_abbr");

        Row row = sheet.createRow(0);
        Cell cell = row.createCell(0);
        cell.setCellValue("id");
        cell = row.createCell(1);
        cell.setCellValue("rid1");
        cell = row.createCell(2);
        cell.setCellValue("rid2");
        cell = row.createCell(3);
        cell.setCellValue("record1");
        cell = row.createCell(4);
        cell.setCellValue("record2");
        cell = row.createCell(5);
        cell.setCellValue("is_duplicate");

        Iterator<String> data1Ie = data.iterator();

        int rownum = 1;
        int qid = 1;
        int record1_id = 1;
        int record2_id = 2;
        while(data1Ie.hasNext()) {
            String record1 = data1Ie.next();
            Iterator<String> data2Ie = data2.iterator();
            while (data2Ie.hasNext()) {
                row = sheet.createRow(rownum);
                String record2 = data2Ie.next();

                cell = row.createCell(0);
                cell.setCellValue(qid);
                qid++;

                cell = row.createCell(1);
                cell.setCellValue(record1_id);
                record1_id += 2;

                cell = row.createCell(2);
                cell.setCellValue(record2_id);
                record2_id += 2;

                cell = row.createCell(3);
                cell.setCellValue(record1);

                cell = row.createCell(4);
                cell.setCellValue(record2);

                cell = row.createCell(5);
                if (Objects.equals(record1, record2)) {
                    cell.setCellValue(1);
                } else {
                    cell.setCellValue(0);
                }

                System.out.println("Records pair " + rownum + " of " + data.size()*data2.size() + " processed.");
                rownum++;
            }
        }

        return workbook;



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

    public static void saveWorkbook(Workbook workbook, String name) {
        try {
            FileOutputStream out = new FileOutputStream(new File("dataset_for_doc2vec_simple_pairs.xlsx"));
            workbook.write(out);
            out.close();
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

}
