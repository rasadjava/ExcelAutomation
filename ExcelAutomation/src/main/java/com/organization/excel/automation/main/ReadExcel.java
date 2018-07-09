package com.organization.excel.automation.main;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.time.Instant;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.Iterator;
import java.util.List;
import java.util.Map;
import java.util.Map.Entry;

import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ReadExcel {
	public static void main(String[] args) {
		String excelPath = "D:\\USAutomation\\excel_sheets\\";
		String sheetName = "Sheet1";
		String outputSheetName = "FilteredExcel1.xls";
		System.out.println("[INFO]: Filtered list " + ReadExcel.comaperExcels(ReadExcel.storeExcelsContent(excelPath, sheetName)));
		ReadExcel.generateNewExcel(ReadExcel.comaperExcels(ReadExcel.storeExcelsContent(excelPath, sheetName)), excelPath, outputSheetName);
	}
	
	public static void generateNewExcel(Map<Integer, ArrayList<String>> finalMapContent, String outputPath, String outputSheetName) {
		try {
            HSSFWorkbook workbook = new HSSFWorkbook();
            HSSFSheet sheet = workbook.createSheet("FirstSheet");  
            for(Entry<Integer, ArrayList<String>> mapContent: finalMapContent.entrySet()) {
            	short excelIndex = mapContent.getKey().shortValue();
            	HSSFRow row = sheet.createRow(excelIndex);
            	List<String> excelListContent = mapContent.getValue();
            	int count= -1;
            	for(String excelContent: excelListContent) {
            		count++;
            		 row.createCell(count).setCellValue(excelContent);
            	}
            }
            FileOutputStream fileOut = new FileOutputStream(outputPath+"/"+outputSheetName);
            workbook.write(fileOut);
            fileOut.close();
            workbook.close();
            System.out.println("Your excel file has been generated!");

        } catch ( Exception ex ) {
            System.out.println(ex);
        }
	}

	public static Map<Integer, ArrayList<String>> comaperExcels(Map<String, Map<Integer, ArrayList<String>>> excelsContant) {
		Map<Integer, ArrayList<String>> excelOne = new HashMap<Integer, ArrayList<String>>();
		Map<Integer, ArrayList<String>> excelTwo = new HashMap<Integer, ArrayList<String>>();
		for (Entry<String, Map<Integer, ArrayList<String>>> excelContent : excelsContant.entrySet()) {
			System.out.println("[INFO]:****" + excelContent);
			if (excelContent.getKey().equals("List of user to compare.xlsx")) {
				excelTwo  = excelContent.getValue();
			} else if (excelContent.getKey().equals("List of users from which records needs to be deleted.xlsx")) {
				excelOne = excelContent.getValue();
			}
		}

		for (Entry<Integer, ArrayList<String>> contentOne : excelOne.entrySet()) {
			List<String> contentOneValues = contentOne.getValue();
			if (contentOne.getKey() > 0) {
				String machineName = contentOneValues.get(0);
				// excelTwo logic
  				excelTwo = ReadExcel.excelTwoLogic(excelTwo, machineName);
			}
		}
		return excelTwo;
	}
	
	public static Map<Integer, ArrayList<String>> excelTwoLogic(Map<Integer, ArrayList<String>> excelTwo, String machineName) {
		Iterator it = excelTwo.entrySet().iterator();
    	while (it.hasNext())
    	{
    	   Entry item = (Entry) it.next();
    	   List<String> list=(List<String>) item.getValue();
    	   if(list.get(0).equals(machineName)) {
    		   it.remove();
    	}
    	   }
		return excelTwo;
	}
	
	public static Map<String, Map<Integer, ArrayList<String>>> storeExcelsContent(String excelPath, String sheetName) {
		String mapKey = null;
		Map<String, Map<Integer, ArrayList<String>>> excelsContent = new HashMap<String, Map<Integer, ArrayList<String>>>();
		Map<Integer, ArrayList<String>> excelContent = new HashMap<Integer, ArrayList<String>>();
		try {
			List<String> listOfExcelFiles = ReadExcel.readAllFileNames(excelPath);
			for (String excelFile : listOfExcelFiles) {
				if (excelFile.contains("List of users from which records needs to be deleted") || excelFile.contains("List of user to compare")) {
					if (excelFile.contains("List of user to compare")) {
						mapKey = excelFile;
					}
					if (excelFile.contains("List of users from which records needs to be deleted")) {
						mapKey = excelFile;
					}
					System.out.println("Copying excel content from the sheet "+mapKey);
					long start = Instant.now().getEpochSecond();
					excelContent = ReadExcel.readExcelContent(excelPath + excelFile, sheetName);
					long end = Instant.now().getEpochSecond();
					System.out.println("Time for " + mapKey + " = " + (start - end) + " seconds");
					excelsContent.put(mapKey, excelContent);
				}
			}
			System.out.println("Copied excel content successfully");
			return excelsContent;
		} catch (IOException e) {
			System.out.println("Error occured while reading excel content");
			e.printStackTrace();
		}
		return excelsContent;
	}

	public static Map<Integer, ArrayList<String>> readExcelContent(String excelFilePath, String sheetName)
			throws IOException {
		int sheetNumber = 0;
		Map<Integer, ArrayList<String>> bunchOfRows = new HashMap<Integer, ArrayList<String>>();
		ArrayList<String> rowList;
		FileInputStream excelInputStream = new FileInputStream(excelFilePath);
		XSSFWorkbook wb = new XSSFWorkbook(excelInputStream);
		if (sheetName.equals("Sheet1")) {
			sheetNumber = 0;
		}
		XSSFSheet sheet = wb.getSheetAt(sheetNumber);
		XSSFRow row;
		XSSFCell column;
		Iterator rows = sheet.rowIterator();
		int count = -1;
		while (rows.hasNext()) { // fetching each row
			count++;
			row = (XSSFRow) rows.next();
			Iterator columns = row.cellIterator();
			rowList = new ArrayList<String>();
			while (columns.hasNext()) { // reading row content one by one
				column = (XSSFCell) columns.next();
				if (column.getCellType() == XSSFCell.CELL_TYPE_STRING) {
					rowList.add(column.getStringCellValue());
				} else if (column.getCellType() == XSSFCell.CELL_TYPE_NUMERIC) {
					rowList.add(String.valueOf((int) column.getNumericCellValue()));
				}
			}
			bunchOfRows.put(count, rowList);
		}
		System.out.println("[INFO]: Reading Excel Sheet '" + sheet.getSheetName() + "'");
		return bunchOfRows;
	}
	
	 public static List<String> readAllFileNames(String directoryPath) throws IOException {
         List<String> allExcelFiles = new ArrayList<String>();
         File dirFile = new File(directoryPath);
         for (File file : dirFile.listFiles()) {
                String excelFile = file.getName();
                if (excelFile.contains(".xlsx") && !(excelFile.contains("~$"))) {
                      allExcelFiles.add(excelFile);
                }
         }
         return allExcelFiles;
  }
}
