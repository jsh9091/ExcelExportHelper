/*
 * Joshua Horvath jsh_197@hotmail.com 
 */

package com.horvath.excelexporthelper;

import java.io.File;
import java.util.ArrayList;
import java.util.List;

/**
 * Main class for EEH library. 
 * @author jhorvath
 */
final public class ExcelExportHelper {

	List<EEHSheet> sheets;
	File file;
	public static final String EXCEPTION_NO_SHEETS_TO_WRITE = "There are no sheets to write to the file.";
	
	/**
	 * Constructor that accepts a filename with a file path as a string.
	 * @param filepath String
	 */
	public ExcelExportHelper(String filepath) {
		this(new File(filepath));
	}
	
	/**
	 * Constructor that accepts a file. 
	 * @param file File 
	 */
	public ExcelExportHelper(File file) {
		
		String fileName = file.getName();
		String parent = file.getParent();
		
		try {
			fileName = FileUtility.validateFileName(fileName);
			FileUtility.testFileLocationWriteable(file);
		} catch (EEHException e) {
			throw new IllegalArgumentException(e.getMessage());
		}

		this.file = new File(parent + File.separator + fileName);
		this.sheets = new ArrayList<>();
	}
	
	/**
	 * Creates and returns an EEHSheet. 
	 * @param sheetName String 
	 * @return EEHSheet
	 */
	public EEHSheet createSheet(String sheetName) {
		EEHSheet sheet = null;
		try {
			sheet = new EEHSheet(sheetName, this.sheets);
			this.sheets.add(sheet);
		} catch (EEHException e) {
			throw new IllegalArgumentException(e.getMessage());
		}
		return sheet;
	}
	
	/**
	 * Triggers operations for writing the Excel file to disk.
	 * @throws EEHException
	 */
	public void writeWorkBook() throws EEHException {
		if (this.sheets.isEmpty()) {
			throw new IllegalStateException(EXCEPTION_NO_SHEETS_TO_WRITE);
		}
		
		EEHExcelFileWriter writer = new EEHExcelFileWriter(this.file, this.sheets);
		writer.writeFile();
	}
	
	public File getFile() {
		return this.file;
	}
	
	public List<EEHSheet> getSheets() {
		return this.sheets;
	}
}
