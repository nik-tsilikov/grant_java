package dataset_preparing;

import com.google.common.base.Strings;
import org.apache.commons.codec.binary.StringUtils;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.*;

public class dataset_name_mtr_with_typos {

    public static final String FILE_PATH_1 = "dataset_name_mtr_dedup.xlsx";
    public static final String FILE_PATH_2 = "univ_dataset_all_words.xlsx";
    public static final String FILE_PATH_3 = "orfo_and_typos.L1_5.xlsx";

    public static void main(String args[]) {
        try {
            TreeSet<String> records = loadExcel(FILE_PATH_1);
            System.out.println("Records list loaded");
            TreeSet<String> existingWords = loadExcel(FILE_PATH_2);
            System.out.println("Existing words list loaded");
            TreeMap<String, HashSet<String>> words_with_orfos = loadMapExcel(FILE_PATH_3, existingWords);
            System.out.println("Words with orfos map loaded");
            saveWorkbook(prepareExcel(getRecordsWithTyposVariations(records, words_with_orfos)));
            System.out.println("Workbook saved");
        } catch (Exception e) {
            System.out.println(e.getMessage());
        }
    }

    private static TreeMap<String, HashSet<String>> getRecordsWithTyposVariations (TreeSet<String> records,TreeMap<String, HashSet<String>> words_with_orfos) {
        TreeMap<String, HashSet<String>> result = new TreeMap<>();
        Set<String> original_words = words_with_orfos.keySet();
        boolean was_replacements = false;
        System.out.println("Records list size: " + records.size());
        for (String record : records) {
            HashSet<String> recordVariations = new HashSet<>();
            recordVariations.add(record);
            String cleanedRecord = cleanRecord(" " + record + " ");
            System.out.println("Record " + record + "cleaned");
            for (String orginal_word : original_words) {
                if (cleanedRecord.contains(" " + orginal_word + " ")) {
                    HashSet<String> replasements = words_with_orfos.get(orginal_word);
                    for (String replacement : replasements) {
                        String tempRecord = record.replace(orginal_word, replacement);
                        System.out.println("Replacement done");
                        was_replacements = true;
                        //result.put(record, tempRecord);
                        recordVariations.add(tempRecord);
                    }
                }
            }
            result.put(record, recordVariations);
        }
//        if (was_replacements) {
//            result.addAll(getRecordsWithTyposVariations(result, words_with_orfos));
//        }
        return result;
    }


    private static TreeMap<String, HashSet<String>> loadMapExcel(String filename, TreeSet<String> existingWords) throws IOException, InvalidFormatException  {
        Cell cell = null;
        int i = 0;
        try {
            TreeMap<String, HashSet<String>> result = new TreeMap<>();

            Workbook workbook = WorkbookFactory.create(new File(filename));

            Sheet sheet = workbook.getSheetAt(0);

            DataFormatter dataFormatter = new DataFormatter();

            Iterator<Row> rowIterator = sheet.rowIterator();


            String prevWord = "";
            HashSet<String> wordVariations = new HashSet<>();
            while (rowIterator.hasNext()) {
                i++;
                Row row = rowIterator.next();
                cell = row.getCell(0);
                if (cell != null) {
                    String currWord = cell.getStringCellValue();
                    if (existingWords.contains(currWord)) {
                        if (!Objects.equals(currWord, prevWord)) {
                            if (wordVariations.size() != 0) {
                                result.put(currWord, wordVariations);
                            }
                            prevWord = currWord;
                            wordVariations = new HashSet<>();
                        }
                        cell = row.getCell(1);
                        if (cell != null) {
                            wordVariations.add(cell.getStringCellValue());
                        }
                    }
                }
            }

            return result;
        } catch (Exception e) {
            System.out.println("Error cell: " + cell + ", row " + i);
            throw e;
        }
    }

    private static TreeSet<String> loadExcel(String filename) throws IOException, InvalidFormatException {

        TreeSet<String> result = new TreeSet<>();

        Workbook workbook = WorkbookFactory.create(new File(filename));

        Sheet sheet = workbook.getSheetAt(0);

        DataFormatter dataFormatter = new DataFormatter();

        Iterator<Row> rowIterator = sheet.rowIterator();
        while (rowIterator.hasNext()) {
            Row row = rowIterator.next();

            Cell cell = row.getCell(0);
            if (cell != null) {
                if (!(cell.getStringCellValue() == null) && !cell.getStringCellValue().isEmpty()) {
                    result.add(Strings.nullToEmpty(cell.getStringCellValue()));
                }
            }
        }

        return result;
    }

    private static Workbook prepareExcel(TreeMap<String, HashSet<String>>  data) {
        XSSFWorkbook workbook = new XSSFWorkbook();

        XSSFSheet sheet = workbook.createSheet("name_mtr_with_typos");

        Row row;
        Cell cell;
        int rownum = 0;
        int recordNum = 0;
        HashSet<String> records = new HashSet<>(data.keySet());
        for (String record : records) {
            HashSet<String> variations = new HashSet<>(data.get(record));
            for (String variation : variations) {
                row = sheet.createRow(rownum);
                cell = row.createCell(0);
                cell.setCellValue(record);
                cell = row.createCell(1);
                cell.setCellValue(variation);
                rownum++;
            }
            recordNum++;
            System.out.println("Record #"+ (recordNum) + " processed.");
        }

        return workbook;



    }

    public static void saveWorkbook(Workbook workbook) {
        try {
            FileOutputStream out = new FileOutputStream(new File("dataset_name_mtr_with_typos.xlsx"));
            workbook.write(out);
            out.close();
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    private static String cleanRecord(String record) {
        record = record.toLowerCase();
        record = record.replaceAll("[0-9A-Za-z;,.#()/*+-]", "");
        record.replace("№", "");
        record = record.replaceAll(" {2,}", " ");
        return record;
    }
}
