package utils;

import java.io.IOException;

public class ExcelUtilsTest {

public static void main(String[] args) throws IOException {

String projectPath = System.getProperty("user.dir");
ExcelUtils excel = new ExcelUtils(projectPath+"\\excelguru\\excelguru.xls.xls.xlsx", "Sheet1"); 

//excel.getRowCount();
//excel.getCellDataString(0, 0);
//excel.getCellDataNumber(1, 1);    
 }

}