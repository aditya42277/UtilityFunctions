package com.excel.functions;

import java.io.IOException;

public class Driver {
    private static String path = ".\\DataFiles\\student-data.xlsx";
    private static String sheetName = "student-data";
    private static ExcelOperations ex = new ExcelOperations(path, sheetName);

    public static void main(String[] args) {
        int rowCount = 0;
        try{
            rowCount = ex.getRowCount();
        }catch (IOException ex){
            System.out.println(ex.getMessage());
            System.out.println("The row count cannot be evaluated");
        }

        System.out.println("The row count is : " + rowCount);


    }
}
