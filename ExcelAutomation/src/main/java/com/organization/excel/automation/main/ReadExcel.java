package com.organization.excel.automation.main;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.time.Instant;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.Iterator;
import java.util.List;
import java.util.Map;
import java.util.Map.Entry;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ReadExcel {
	public static void main(String[] args) {
		String excelPath = "D:\\USAutomation\\excel_sheets\\";
		String sheetName = "Sheet1";
		System.out.println("[INFO]: Filtered list "+ReadExcel.comaperExcels(ReadExcel.storeExcelsContent(excelPath, sheetName)));
	}

	public static Map<Integer, ArrayList<String>> comaperExcels(
			Map<String, Map<Integer, ArrayList<String>>> excelsContant) {
		Map<Integer, ArrayList<String>> excelOne = new HashMap<Integer, ArrayList<String>>();
		Map<Integer, ArrayList<String>> excelTwo = new HashMap<Integer, ArrayList<String>>();
		for (Entry<String, Map<Integer, ArrayList<String>>> excelContent : excelsContant.entrySet()) {
			System.out.println("[INFO]:****" + excelContent);
			if (excelContent.getKey().equals("List of user to compare.xlsx")) {
				excelOne = excelContent.getValue();
			} else if (excelContent.getKey().equals("List of users from which records needs to be deleted.xlsx")) {
				excelTwo = excelContent.getValue();
			}
		}

		for (Entry<Integer, ArrayList<String>> contentOne : excelOne.entrySet()) {
			List<String> contentOneValues = contentOne.getValue();
			if (contentOne.getKey() > 0) {
				String machineName = contentOneValues.get(0);
				// excelTwo logic
				for (Entry<Integer, ArrayList<String>> contentTwo : excelTwo.entrySet()) {
					List<String> contentTwoValues = contentTwo.getValue();
					if (contentTwo.getKey() > 0) {
						if (contentTwoValues.get(0).equals(machineName)) {
							excelTwo.remove(contentTwo.getKey());
						}
					}
				}
			}
		}
		return excelTwo;
	}
	
	public static Map<String, Map<Integer, ArrayList<String>>> storeExcelsContent(String excelPath, String sheetName) {
		// String excelPath = "D:\\USAutomation\\excel_sheets\\Deployed Machine.xlsx";
		/*String excelPath = "D:\\USAutomation\\excel_sheets\\";
		String sheetName = "Sheet1";*/
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
					// long start = System.currentTimeMillis();
					long start = Instant.now().getEpochSecond();
					excelContent = ReadExcel.readExcelContent(excelPath + excelFile, sheetName);
					// long end = System.currentTimeMillis();
					long end = Instant.now().getEpochSecond();
					System.out.println("Time for " + mapKey + " = " + (start - end));
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
		// SheetContent sheetContent = new SheetContent();
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
		// sheetContent.setSheetName(sheet.getSheetName());
		System.out.println("[INFO]: Reading Excel Sheet '" + sheet.getSheetName() + "'");
		// sheetContent.setSheetContent(bunchOfRows);
		return bunchOfRows;
	}
	
	 public static List<String> readAllFileNames(String directoryPath) throws IOException {
         List<String> allExcelFiles = new ArrayList<String>();
         File dirFile = new File(directoryPath);
         for (File file : dirFile.listFiles()) {
                String excelFile = file.getName();
                if (excelFile.contains(".xlsx") && !(excelFile.contains("~$"))) {
                      allExcelFiles.add(excelFile);
//                    System.out.println("[INFO]: Excel files path> " + excelFile);
                }
         }
         return allExcelFiles;
  }
}
