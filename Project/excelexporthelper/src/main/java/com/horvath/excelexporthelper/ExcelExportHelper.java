/*
 * MIT License
 * 
 * Copyright (c) 2021 Joshua Horvath
 * 
 * Permission is hereby granted, free of charge, to any person obtaining a copy
 * of this software and associated documentation files (the "Software"), to deal
 * in the Software without restriction, including without limitation the rights
 * to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
 * copies of the Software, and to permit persons to whom the Software is
 * furnished to do so, subject to the following conditions:
 * 
 * The above copyright notice and this permission notice shall be included in all
 * copies or substantial portions of the Software.
 * 
 * THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
 * IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
 * FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
 * AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
 * LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
 * OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
 * SOFTWARE.
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

	private List<EEHSheet> sheets;
	private File file;
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
		} catch (EEHException ex) {
			throw new IllegalArgumentException(ex.getMessage(), ex);
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
		} catch (EEHException ex) {
			throw new IllegalArgumentException(ex.getMessage(), ex);
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

	@Override
	public String toString() {
		return "ExcelExportHelper [sheets=" + sheets + ", file=" + file + "]";
	}
	
}
