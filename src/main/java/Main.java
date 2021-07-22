import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.Arrays;
import java.util.HashMap;
import java.util.Map;

public class Main {
    public static void main(String[] args) throws IOException {
        XSSFSheet sheet = readSheetFromExcel("src/main/resources/words.xlsx");
        Map<String, String[]> wordPairs = readWordPairsFromExcelSheet(sheet);
        wordPairs.forEach((k,v) -> System.out.println(k + ' ' + Arrays.toString(v)));
    }

    public static XSSFSheet readSheetFromExcel(String file) throws IOException {
        XSSFWorkbook myExcelBook = new XSSFWorkbook(new FileInputStream(file));
        XSSFSheet sheet = myExcelBook.getSheet("Sheet1");
        myExcelBook.close();
        return sheet;
    }

    public static Map<String, String[]> readWordPairsFromExcelSheet(XSSFSheet excelSheet) {
        Map<String, String[]> wordsPair = new HashMap<>();
        int rowIndex = 0;
        XSSFRow row;
        while ((row = excelSheet.getRow(rowIndex++)) != null) {
            if(row.getCell(0).getCellType() == CellType.STRING &&
                    row.getCell(1).getCellType() == CellType.STRING) {
                String engWord = row.getCell(0).getStringCellValue();
                String[] rusWords = row.getCell(1).getStringCellValue().split("/");
                wordsPair.put(engWord, rusWords);
            }
        }
        return wordsPair;
    }
}
