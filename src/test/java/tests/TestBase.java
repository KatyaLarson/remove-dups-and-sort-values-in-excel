package tests;

import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.AfterEach;
import org.junit.jupiter.api.BeforeEach;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

public class TestBase {
    protected static final String INPUT_FILE_PATH = "src/test/resources/SampleExcel.xlsx";
    protected static final String OUTPUT_FILE_PATH = "src/test/resources/SampleExcel.xlsx";
    protected static final String SHEET_NAME = "Names";

    protected static final String NEW_SHEET_NAME = "NoDupsSortedData";

    protected static DataFormatter dataFormatter;
    protected XSSFWorkbook workbook;

    @BeforeEach
    void setUp() {
        dataFormatter = new DataFormatter();
        try {
            FileInputStream fileInputStream = new FileInputStream(INPUT_FILE_PATH);
            workbook = new XSSFWorkbook(fileInputStream);
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    @AfterEach
    void tearDown() {
        try {
            FileOutputStream out = new FileOutputStream(OUTPUT_FILE_PATH);
            workbook.write(out);
            out.close();
            workbook.close();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}
