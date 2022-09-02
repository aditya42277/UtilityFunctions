package com.excel.functions;

import org.apache.poi.xssf.usermodel.*;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;


public class ExcelOperations {

    private String path = null;
    private String sheetName = null;
    private FileInputStream inputStream;
    private FileOutputStream outputStream;
    private XSSFWorkbook workbook;
    private XSSFSheet sheet;
    private XSSFRow row;
    private XSSFCell cell;
    private XSSFCellStyle cellStyle;

    public ExcelOperations(String path, String sheetName) {
        this.path = path;
        this.sheetName = sheetName;
    }

    //Method to find the row count in an Excel sheet
    public int getRowCount() throws IOException {
        inputStream = new FileInputStream(path);
        workbook = new XSSFWorkbook(inputStream);
        if (workbook == null) {
            return -99;
        }
        sheet = workbook.getSheet(sheetName);
        if(sheet == null){
            return -99;
        }
        int totalRowCount = sheet.getLastRowNum();
        return totalRowCount;

    }


}
