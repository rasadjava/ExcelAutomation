package com.organization.excel.automation.main;

//Find jar from here "http://poi.apache.org/download.html"
import  java.io.*;
import  org.apache.poi.hssf.usermodel.HSSFSheet;
import  org.apache.poi.hssf.usermodel.HSSFWorkbook;
import  org.apache.poi.hssf.usermodel.HSSFRow;

public class GenerateNewExcel {


	    public static void main(String[]args) {
	        try {
	            String filename = "D:/NewExcelFile.xls" ;
	            HSSFWorkbook workbook = new HSSFWorkbook();
	            HSSFSheet sheet = workbook.createSheet("FirstSheet");  

	            HSSFRow rowhead = sheet.createRow((short)0);
	            rowhead.createCell(0).setCellValue("No.");
	            rowhead.createCell(1).setCellValue("Name");
	            rowhead.createCell(2).setCellValue("Address");
	            rowhead.createCell(3).setCellValue("Email");

	            HSSFRow row = sheet.createRow((short)2);
	            row.createCell(0).setCellValue("1");
	            row.createCell(1).setCellValue("Sankumarsingh");
	            row.createCell(2).setCellValue("India");
	            row.createCell(3).setCellValue("sankumarsingh@gmail.com");

	            FileOutputStream fileOut = new FileOutputStream(filename);
	            workbook.write(fileOut);
	            fileOut.close();
	            workbook.close();
	            System.out.println("Your excel file has been generated!");

	        } catch ( Exception ex ) {
	            System.out.println(ex);
	        }
	    }


}
