package com.excel.functions;

import org.apache.poi.xssf.usermodel.*;

import java.io.FileInputStream;
import java.io.FileOutputStream;


public class ExcelOperations {

    private String path = null;
    private FileInputStream inputStream;
    private FileOutputStream outputStream;
    private XSSFWorkbook workbook;
    private XSSFSheet sheet;
    private XSSFRow row;
    private XSSFCell cell;
    private XSSFCellStyle cellStyle;

    public ExcelOperations(String path) {
        this.path = path;
    }

    
}
