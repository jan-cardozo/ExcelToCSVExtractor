package com.netpeople.tools;

import java.io.BufferedWriter;
import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStreamWriter;
import java.text.DateFormat;
import java.util.ArrayList;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class ExcelToCSVExtractor {
	
	private static int headerLength = 0;
	private static int realHL = 0;

	private static String join(ArrayList<String> array, String delimiter) {
		StringBuffer s = new StringBuffer();
		for (int i = 0; i < array.size(); i++) {
			s.append(array.get(i));
			if (i < (array.size() - 1)) {
				s.append(delimiter);
			} else {
				s.append(System.getProperty("line.separator"));
			}
		}
		return s.toString();
	}

	/**
	 * @param args
	 */
	public static void main(String[] args) {
		try {
			File in = new File(args[0]);
			if (!in.canRead()) {
				throw new FileNotFoundException();
			}

			File out = new File(args[1]);
			if (!out.createNewFile()) {
				throw new FileNotFoundException();
			}			
						
			BufferedWriter writer = new BufferedWriter(
					new OutputStreamWriter(
							new FileOutputStream(out), "UTF-8"));

			Workbook workbook = WorkbookFactory.create(in);

			Sheet sheet = workbook.getSheetAt(0);
			
			headerLength = sheet.getRow(0).getLastCellNum();
			
			int headerStart = sheet.getRow(0).getFirstCellNum();
		
			for (int cellNumber = headerStart; cellNumber < headerLength; cellNumber++) {
				Cell cell = sheet.getRow(0).getCell(cellNumber);
				if (cell != null) {
					if (cell.getStringCellValue().length() > 0) {
						realHL++;
					}else{
						break;
					}
				} else{
					realHL++;
				}
			}
			
			for (Row row : sheet) {
				ArrayList<String> array = new ArrayList<String>();
				for (int cellNumber = headerStart; cellNumber < realHL; cellNumber++) {
					
					if(cellNumber >= 0){
						Cell cell = row.getCell(cellNumber);
						if (cell != null) {
							switch (cell.getCellType()) {
							case Cell.CELL_TYPE_BLANK:
								array.add("");
								break;
	
							case Cell.CELL_TYPE_BOOLEAN:
								array.add(String.valueOf(cell.getBooleanCellValue()));
								break;
	
							case Cell.CELL_TYPE_ERROR:
	
								array.add("");
								break;
	
							case Cell.CELL_TYPE_FORMULA:
								switch (cell.getCachedFormulaResultType()) {
								case Cell.CELL_TYPE_NUMERIC:
									array.add(String.valueOf(cell.getNumericCellValue()));
									break;
	
								case Cell.CELL_TYPE_STRING:
									array.add(cell.getStringCellValue().replaceAll("[\n\r|]", "").trim());
	
								default:
									break;
								}
								break;
	
							case Cell.CELL_TYPE_NUMERIC:
								if (DateUtil.isCellDateFormatted(cell)) {
									array.add(DateFormat.getDateInstance().format(
											cell.getDateCellValue()));
								} else {
									array.add(String.valueOf(cell.getNumericCellValue()));
								}
								break;
	
							case Cell.CELL_TYPE_STRING:
								array.add(cell.getStringCellValue().replaceAll("[\n\r|]", "").trim());
								break;
	
							default:
								System.out.println("default");
								array.add("");
								break;
							}
						} else {
							array.add("");
						}
					}
				}
				
				boolean emptyRow = true;
				for (String string : array) {
					if(!string.isEmpty()){
						emptyRow = false;
					}
				}

				if(!emptyRow){
					writer.write(join(array, "|"));
				}
			}
			writer.close();
		} catch (FileNotFoundException e) {
			System.err
					.println("There was a problem while accessing one of the files1");
			System.exit(1);
		} catch (InvalidFormatException e) {
			System.err
					.println("There was a problem while reading the excel file2");
			System.exit(2);
		} catch (IOException e) {
			System.err
					.println("There was a problem while reading the excel file3");
			System.exit(2);
		} catch (Exception e) {
			System.err.println(e.getMessage());
			System.exit(3);
		}
		System.exit(0);
	}
}
