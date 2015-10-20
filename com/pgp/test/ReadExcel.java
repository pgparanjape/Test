package com.pgp.test;

import java.io.FileInputStream;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.List;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class ReadExcel {

//Updating
	Workbook wbWorkbook;
	Sheet shSheet;
	static String filePath= "D:\\TestCases.xls";

	public static void main(String[] args) {
		List<String> myList = new ArrayList<String>();
		

		try {
			InputStream inp = new FileInputStream(filePath);
			Workbook wb = null;
			wb = WorkbookFactory.create(new File(filePath));
			Sheet sheet = wb.getSheet("Execution");
			for (int i = 1; i<=sheet.getLastRowNum(); i++){
				Row row = sheet.getRow(i);
				Cell testCaseSheetNameCell = row.getCell(1);
				Cell testCaseSheetExecutionCell = row.getCell(2);
				if (testCaseSheetExecutionCell != null && testCaseSheetNameCell != null){
					if (testCaseSheetNameCell.getCellType() == Cell.CELL_TYPE_STRING && testCaseSheetExecutionCell.getCellType() == Cell.CELL_TYPE_STRING){
						System.out.println(testCaseSheetNameCell.getStringCellValue() + "   " +testCaseSheetExecutionCell.getStringCellValue());
						myList.add(testCaseSheetNameCell.getStringCellValue()+"@"+testCaseSheetExecutionCell.getStringCellValue());
					}
				}
			}

			for (String sheetName : myList) {
				String [] SheetNameValues = sheetName.split("@");
				if (SheetNameValues[1].equalsIgnoreCase("Y")){

					boolean sheetExists;
					int index = wb.getSheetIndex(SheetNameValues[0]);
					if(index==-1){
						index=wb.getSheetIndex(sheetName.toUpperCase());
						if(index==-1)
							sheetExists =  false;
						else
							sheetExists=  true;
					}
					else
						sheetExists =  true;

					if (sheetExists){ 
						sheet = wb.getSheet(SheetNameValues[0]);
						for (int row = 1; row <= sheet.getLastRowNum(); row++) {
							if (sheet.getRow(row).getCell(0) != null){
								System.out.println(sheet.getRow(row).getCell(0));
							}
							for (int col = 1; col <= sheet.getRow(col).getLastCellNum(); col++) {
								if(sheet.getRow(row).getCell(col) != null){
									if (sheet.getRow(row).getCell(col).getCellType() == Cell.CELL_TYPE_STRING){
										String params = sheet.getRow(row).getCell(col).getStringCellValue();
										String [] paramsSplit = params.split(":");
										if (paramsSplit.length == 2) {
											System.out.println(paramsSplit[0] + " : " + paramsSplit[1]);
										}

									}
								}

							}

						}
					}
				}
			}
		
		} catch (FileNotFoundException e) {
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		}catch (InvalidFormatException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
	}
}
