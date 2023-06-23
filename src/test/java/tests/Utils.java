package tests;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.util.*;

public class Utils extends TestBase {



    public  static void removeSheet(String sheetName) throws IOException {

       FileInputStream inputStream=new FileInputStream(new File("src/test/resources/SampleExcel.xlsx"));
       XSSFWorkbook workbook=new XSSFWorkbook(inputStream);

        for(int i=0;i< workbook.getNumberOfSheets();i++){
            XSSFSheet workSheet= workbook.getSheetAt(i);
            if(workSheet.getSheetName().equals(sheetName)){
                workbook.removeSheetAt(i);
            }
        }

       FileOutputStream outFile=new FileOutputStream(new File("src/test/resources/SampleExcel.xlsx"));
        workbook.write(outFile);
        outFile.close();
    }

    public static void removeSheet(Workbook workbook, String sheetName) {
        int sheetIndex = workbook.getSheetIndex(sheetName);
        if (sheetIndex >= 0) {
            workbook.removeSheetAt(sheetIndex);
        }
    }


    public static String getColumnsNames(XSSFSheet workSheet) {

        String str="";
        for (Cell cell : workSheet.getRow(0)) {
            str+=cell.toString()+",";
        }
        return str;
    }

    public static void writeUniqueRowsToSheet(Set<String> set,XSSFSheet sheet,String headers) {
        Map<String, Object> data = new TreeMap<>();


        List<String> list = new ArrayList<>(set);
        list.add(0, headers);

        for (int i = 0; i < list.size(); i++) {
            String[] arr = list.get(i).split(",");
            data.put("\"" + i + "\"", arr);
        }

        int rowId = 0;
        for (String key : data.keySet()) {
            Row newRow = sheet.createRow(rowId++);
            Object[] objectArr = (Object[]) data.get(key);
            int cellId = 0;

            for (Object obj : objectArr) {
                Cell cell = newRow.createCell(cellId++);
                cell.setCellValue((String) obj);
            }
        }
    }


    public static Set<String> getUniqueRows(XSSFSheet sheet) {
        Set<String> set = new TreeSet<>(String.CASE_INSENSITIVE_ORDER);

        Iterator<Row> rows=sheet.rowIterator();

        while (rows.hasNext()) {
            Row row= rows.next();
            if (row.getRowNum()==0){
                continue;
            }
            StringBuilder oneLine = new StringBuilder();
            for (Cell cell : row) {
                String cellValueStr = dataFormatter.formatCellValue(cell);
                oneLine.append(cellValueStr).append(",");
            }
            if (oneLine.length() > 0) {
                oneLine.setLength(oneLine.length() - 1);
            }
            set.add(oneLine.toString().trim());
        }


        return set;
    }


}
