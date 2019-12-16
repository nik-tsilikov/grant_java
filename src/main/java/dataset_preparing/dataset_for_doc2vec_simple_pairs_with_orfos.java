package dataset_preparing;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
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

            //saveWorkbook(prepareExcel(loadExcel(FILE_PATH)));
            prepareExcel(loadExcel(FILE_PATH));

        } catch (Exception e) {
            System.out.println(e.getMessage());
        }
    }

    private static void prepareExcel(HashMap<String, String> data) {
        HashSet<String> allRecords1 = new HashSet<>();
        allRecords1.addAll(data.keySet());

        HashSet<String> allRecords2 = new HashSet<>(allRecords1);
      //  XSSFWorkbook workbook = new XSSFWorkbook();

       // XSSFSheet sheet = workbook.createSheet("doc2vec_test_dataset");

        SXSSFWorkbook wb = null;
        FileOutputStream fos = null;
        try {
            wb = new SXSSFWorkbook(100);//SXSSFWorkbook.DEFAULT_WINDOW_SIZE/* 100 */);
            Sheet sheet = wb.createSheet();
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
            int file_index = 1;
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
                    if (Objects.equals(data.get(record1), data.get(record2))) {
                        cell.setCellValue(1);
                    } else {
                        cell.setCellValue(0);
                    }

                    System.out.println("Records pair " + (rownum + (file_index-1)*1000000) + " of " + (allRecords1.size()*allRecords2.size()) + " processed.");

                    if (rownum % 1000000 == 0) {
                        fos = new FileOutputStream(new File("dataset_for_doc2vec_simple_pairs_with_typos_" + file_index + ".xlsx"));
                        file_index++;
                        wb.write(fos);
                        wb = new SXSSFWorkbook(100);//SXSSFWorkbook.DEFAULT_WINDOW_SIZE/* 100 */);
                        sheet = wb.createSheet();
                        row = sheet.createRow(0);
                        cell = row.createCell(0);
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
                        rownum = 1;
                    }
                    rownum++;

                }
            }
            fos = new FileOutputStream(new File("dataset_for_doc2vec_simple_pairs_with_typos_" + file_index + ".xlsx"));
            file_index++;
            wb.write(fos);
        } catch (Exception e) {
            System.out.println(e.getMessage());
        }finally {
            try {
                if (fos != null) {
                    fos.close();
                }
            } catch (IOException e) {
            }
            try {
                if (wb != null) {
                    wb.close();
                }
            } catch (IOException e) {
            }
        }



//        return workbook;



    }

    private static HashMap<String, String> loadExcel(String filename) throws IOException, InvalidFormatException {

        HashMap<String, String> result = new HashMap<>();

        Workbook workbook = WorkbookFactory.create(new File(filename));

        Sheet sheet = workbook.getSheetAt(0);

        DataFormatter dataFormatter = new DataFormatter();

        Iterator<Row> rowIterator = sheet.rowIterator();


        while (rowIterator.hasNext()) {
            Row row = rowIterator.next();
            Cell cell = row.getCell(0);
            String baseRecord = cell.getStringCellValue();
            cell = row.getCell(1);
            String changedRecord = cell.getStringCellValue();
            result.put(changedRecord, baseRecord);
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

//    public static String getRecordKey(String record, HashMap<String, HashSet<String>> allRecords) {
//        final String[] result = new String[1];
//        allRecords.forEach((records, recordsVars) -> {
//            if (recordsVars.contains(record)) {
//                result[0] = record;
//            }
//        });
//        return result[0];
//    }

}
