package com.tariq.readexcelwritexml;

import java.io.BufferedInputStream;
import java.io.BufferedWriter;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileWriter;
import java.io.IOException;
import java.io.InputStream;
import java.io.Writer;
import java.util.Iterator;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;

public class ReadExcelWriteToXML {

	public static void main(String[] args) {
		Writer dynamicWriter = null;
		try {

			InputStream input = new BufferedInputStream(new FileInputStream(
					"/home/bjit-8/Documents/workspace_tariq/ReadFromExcelWriteToXML/resource/Sample.xls"));
			POIFSFileSystem fs = new POIFSFileSystem(input);
			HSSFWorkbook wb = new HSSFWorkbook(fs); // Read Workbook from
													// filesystem

			for (int sheetIndex = 1; sheetIndex < wb.getNumberOfSheets() - 1; sheetIndex++) {
				HSSFSheet sheet = wb.getSheetAt(sheetIndex); // sheet of excel
				Iterator rows = sheet.rowIterator(); // Get Rows of the specific
														// sheet
				File dynamicFile = null;
				String root = "Responses/Actual/";
				if (new File(root + sheet.getSheetName()).mkdirs()) {
					System.out.println("Root directory created successfully!");
				}
				while (rows.hasNext()) {
					HSSFRow row = (HSSFRow) rows.next();
					System.out.println("\n");
					if (row.getRowNum() > 1) {
						Iterator cells = row.cellIterator(); // Get all cells
																// from current
																// row
						while (cells.hasNext()) {
							HSSFCell cell = (HSSFCell) cells.next();
							if (cell.getColumnIndex() == 0) {
								if ((cell.getCellType() == HSSFCell.CELL_TYPE_STRING)
										|| (cell.getCellType() == HSSFCell.CELL_TYPE_NUMERIC)) {
									dynamicFile = null;
									if (sheetIndex != 0) {
										if (!new File(root + sheet.getSheetName()).exists()) { // Check
											// if
											// Directory
											// already
											// exists
											System.out.println(new File(root + sheet.getSheetName()).mkdir()); // Create
											// Directory
										}
										if (sheetIndex < 10) {
											dynamicFile = new File(root + sheet.getSheetName() + "/"
													+ sheet.getSheetName() + "-Normal-"
													+ String.valueOf(cell.getStringCellValue()) + ".xml");
										} else if (sheetIndex == 10) {
											dynamicFile = new File(root + sheet.getSheetName() + "/" + "Author-Case-"
													+ String.valueOf(cell.getStringCellValue()) + ".xml");
										} else if (sheetIndex == 11) {
											dynamicFile = new File(root + sheet.getSheetName() + "/" + "Common-Error-"
													+ String.valueOf(cell.getStringCellValue()) + ".xml");
										}
									} else {
										System.out.println("Invalid Sheet Selection");
									}

								} else {
									System.out.print("Unexpected cell type");
								}
							}

							if ((sheetIndex < 10 && cell.getColumnIndex() == 12)
									|| (sheetIndex == 10 && cell.getColumnIndex() == 13)
									|| (sheetIndex == 11 && cell.getColumnIndex() == 14)) {

								if ((cell.getCellType() == HSSFCell.CELL_TYPE_STRING)
										|| (cell.getCellType() == HSSFCell.CELL_TYPE_BLANK)) {
									System.out.println(dynamicFile);
									System.err.println(dynamicFile != null ? dynamicFile.getName() : "");
									System.err.println("Actual Response: " + cell.getStringCellValue());
									if (dynamicFile != null) {
										dynamicWriter = new BufferedWriter(new FileWriter(dynamicFile));
										dynamicWriter.write(cell.getStringCellValue());
										dynamicFile = null;
										dynamicWriter.close();
									}
								} else {
									System.out.print("Unexpected cell type");
								}
							}
						}
					}
				}
			}
		} catch (IOException ex) {
			ex.printStackTrace();
		} finally {
			try {
				if (dynamicWriter != null) {
					dynamicWriter.close();
				}
			} catch (IOException e) {
				e.printStackTrace();
			}
		}
	}

}
