package dataset_preparing;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.HashMap;
import java.util.HashSet;
import java.util.Iterator;
import java.util.Objects;

public class dataset_for_doc2vec_simple_pairs_with_orfos {

    public static final String FILE_PATH = "name_mtr_with_typos.xlsx";

    public static void main(String args[]) {
        try {

            saveWorkbook(prepareExcel(loadExcel(FILE_PATH)));

        } catch (Exception e) {
            System.out.println(e.getMessage());
        }
    }

    private static Workbook prepareExcel(HashMap<String, HashSet<String>> data) {
        HashSet<String> allRecords1 = new HashSet<>();
        data.forEach((s, strings) -> allRecords1.addAll(strings));

        HashSet<String> allRecords2 = new HashSet<>(allRecords1);
        XSSFWorkbook workbook = new XSSFWorkbook();

        XSSFSheet sheet = workbook.createSheet("doc2vec_test_dataset");

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

        Iterator<String> data1Ie = allRecords1.iterator();

        int rownum = 1;
        int qid = 1;
        int record1_id = 1;
        int record2_id = 2;
        while(data1Ie.hasNext()) {
            String record1 = data1Ie.next();
            Iterator<String> data2Ie = allRecords2.iterator();
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
                if (Objects.equals(getRecordKey(record1, data), getRecordKey(record2,data))) {
                    cell.setCellValue(1);
                } else {
                    cell.setCellValue(0);
                }

                System.out.println("Records pair " + rownum + " of " + allRecords1.size()*allRecords2.size() + " processed.");
                rownum++;
            }
        }

        return workbook;



    }

    private static HashMap<String, HashSet<String>> loadExcel(String filename) throws IOException, InvalidFormatException {

        HashMap<String, HashSet<String>> result = new HashMap<String, HashSet<String>> ();

        Workbook workbook = WorkbookFactory.create(new File(filename));

        Sheet sheet = workbook.getSheetAt(0);

        DataFormatter dataFormatter = new DataFormatter();

        Iterator<Row> rowIterator = sheet.rowIterator();

        String prevRecord = "";
        HashSet<String> recordsVariations = new HashSet<>();
        while (rowIterator.hasNext()) {
            Row row = rowIterator.next();
            Cell cell = row.getCell(0);
            String currRecord = cell.getStringCellValue();
            if (!Objects.equals(currRecord, prevRecord)) {
                if (recordsVariations.size() != 0) {
                    result.put(prevRecord, recordsVariations);
                }
                prevRecord = currRecord;
                recordsVariations = new HashSet<>();
            }
            cell = row.getCell(1);
            if (cell != null) {
                recordsVariations.add(cell.getStringCellValue());
            }
        }

        return result;
    }

    public static void saveWorkbook(Workbook workbook) {
        try {
            FileOutputStream out = new FileOutputStream(new File("dataset_for_doc2vec_simple_pairs_with_typos.xlsx"));
            workbook.write(out);
            out.close();
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    public static String getRecordKey(String record, HashMap<String, HashSet<String>> allRecords) {
        final String[] result = new String[1];
        allRecords.forEach((records, recordsVars) -> {
            if (recordsVars.contains(record)) {
                result[0] = record;
            }
        });
        return result[0];
    }

}
