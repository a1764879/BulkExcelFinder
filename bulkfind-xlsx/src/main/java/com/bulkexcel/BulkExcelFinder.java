package com.bulkexcel;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashSet;
import java.util.Iterator;
import java.util.LinkedList;
import java.util.List;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class BulkExcelFinder {
	private static final String DIRECTORY = "C:\\Users\\615429\\Downloads\\RTMs";
	private static final String FILE_NAME = "D:\\MyFirstExcel.xlsx";
	private static ArrayList<String> skipValues = new ArrayList<String>();
	private static List<String> filenames = new LinkedList<String>();
	private static List<Cell> cells = new ArrayList<Cell>();

	public static void main(String[] args) throws Exception {
		XSSFWorkbook workbookWrite = new XSSFWorkbook();
        XSSFSheet sheetWrite = workbookWrite.createSheet("RTM_LIST");
        int rowNumWrite = 0;
		skipValues.add("Coding");
		skipValues.add("Program / Source Code File Name(s)");
		skipValues.add("Not Applicable");
		skipValues.add("Not Applicable. ");
		
		final File folder = new File(DIRECTORY);
		listFilesForFolder(folder);
		System.out.println(filenames.size());

		try {
			for (String fName : filenames) {

				FileInputStream excelFile = new FileInputStream(new File(DIRECTORY+"\\" + fName));
				Workbook workbook = new XSSFWorkbook(excelFile);
				Sheet datatypeSheet = workbook.getSheetAt(0);
				Iterator<Row> iterator = datatypeSheet.iterator();

				String columnWanted = "Program / Source Code File Name(s)";
				Row firstRow = datatypeSheet.getRow(2);
				Integer columnNo = null;
				
				
				

				for (Cell cell : firstRow) {
					if (cell.getStringCellValue().equals(columnWanted)) {
						columnNo = cell.getColumnIndex();
						break;
					}
				}
				System.out.println("----"+fName+"----");
				List<Cell> tempCells = new ArrayList<Cell>();
				if (columnNo != null) {
					for (Row row : datatypeSheet) {
						Cell c = row.getCell(columnNo);
						if (c == null || c.getCellType() == Cell.CELL_TYPE_BLANK
								||skipValues.contains(c.getStringCellValue())) {
							// Nothing in the cell in this row, skip it
						} else {
							Row rowWrite = sheetWrite.createRow(rowNumWrite++);
					        Cell cellWrite = rowWrite.createCell(0);
					        cellWrite.setCellValue((String) fName);
					        cellWrite = rowWrite.createCell(1);
					        cellWrite.setCellValue(c.getStringCellValue());
					        
							cells.add(c);
							tempCells.add(c);
						}
					}
				} else {
					System.out.println("could not find column " + columnWanted);
				}
				
				
				for (Cell cc : tempCells) {
					
					System.out.println(cc.getStringCellValue());
				}
				
				

			}
		} catch (Exception e) {
			e.printStackTrace();
		}
		System.out.println("*************************************************");
		HashSet<String> hset = new HashSet<String>();
		for (Cell cc : cells) {
			hset.add(cc.getStringCellValue());
			System.out.println(cc.getStringCellValue());
		}
		System.out.println(hset.size());
		
		try {
            FileOutputStream outputStream = new FileOutputStream(FILE_NAME);
            workbookWrite.write(outputStream);
            workbookWrite.close();
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }

	}

	public static void listFilesForFolder(final File folder) {
		for (final File fileEntry : folder.listFiles()) {
			if (fileEntry.isDirectory()) {
				listFilesForFolder(fileEntry);
			} else {
				if (fileEntry.getName().contains(".xlsx"))
					filenames.add(fileEntry.getName());
			}
		}
	}


}
