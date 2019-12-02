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
            TreeSet<String> words = loadExcel(FILE_PATH_2);
            TreeMap<String, HashSet<String>> words_with_orfos = loadMapExcel(FILE_PATH_3);
        } catch (Exception e) {
            System.out.println(e.getMessage());
        }
    }

    private static TreeMap<String, HashSet<String>> loadMapExcel(String filename) throws IOException, InvalidFormatException  {
        TreeMap<String, HashSet<String>> result = new TreeMap<>();

        Workbook workbook = WorkbookFactory.create(new File(filename));

        Sheet sheet = workbook.getSheetAt(0);

        DataFormatter dataFormatter = new DataFormatter();

        Iterator<Row> rowIterator = sheet.rowIterator();

        Cell cell;
        String prevWord = "";
        HashSet<String> wordVariations = new HashSet<>();
        while (rowIterator.hasNext()) {
            Row row = rowIterator.next();
            cell = row.getCell(0);
            String currWord = cell.getStringCellValue();
            if (!Objects.equals(currWord, prevWord)) {
                if (wordVariations.size() != 0) {
                    result.put(currWord, wordVariations);
                }
                prevWord = new String(currWord);
                wordVariations = new HashSet<>();
            }
            cell = row.getCell(1);
            wordVariations.add(cell.getStringCellValue());
        }

        return result;
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
            result.add(Strings.nullToEmpty(cell.getStringCellValue()));
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
