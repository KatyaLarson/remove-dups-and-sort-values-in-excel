package tests;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.*;
import tests.Utils;

import java.io.*;
import java.util.*;

@TestMethodOrder(MethodOrderer.OrderAnnotation.class)
class TestExcel extends TestBase {

    @Test
    @Order(1)
    void readFromExcelFile() {
        XSSFSheet sheet = workbook.getSheet(SHEET_NAME);
        if (sheet == null) {
            throw new IllegalArgumentException("Sheet '" + SHEET_NAME + "' not found in the workbook.");
        }

        Set<String> set = Utils.getUniqueRows(sheet);
       // System.out.println(set);

        XSSFSheet newSheet = workbook.createSheet(NEW_SHEET_NAME);
        String headers=Utils.getColumnsNames(sheet);
        Utils.writeUniqueRowsToSheet(set, newSheet,headers);
    }

   /* @Test
    @Order(2)
    void deleteSheet() throws IOException {
        Utils.removeSheet(workbook,NEW_SHEET_NAME);
    }*/
}


