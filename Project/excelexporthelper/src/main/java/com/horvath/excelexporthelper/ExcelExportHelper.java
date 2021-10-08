/*
 * Joshua Horvath jsh_197@hotmail.com 
 */

package com.horvath.excelexporthelper;

import java.io.File;
import java.util.List;

final public class ExcelExportHelper {

	List<EEHSheet> sheets;
	File file;
	
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
	}
	
	public File getFile() {
		return this.file;
	}
}
