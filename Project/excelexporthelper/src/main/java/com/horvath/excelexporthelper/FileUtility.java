/*
 * Joshua Horvath jsh_197@hotmail.com 
 */

package com.horvath.excelexporthelper;

import java.io.File;

final class FileUtility {
	
	public static final String EXCEL_SUFFIX = ".xlsx";
	
	/**
	 * Ensures that the name of the file is valid.
	 * @param filename String 
	 * @return String 
	 * @throws EEHException 
	 */
	protected static String validateFileName(String filename) throws EEHException {
		
		if (filename.isEmpty()) {
			throw new EEHException("File name must not be empty.");
		}			
		
		filename = checkFileNameLength(filename);
		
		filename = checkFileNameSuffix(filename);
		
		return filename;
	}

	/**
	 * Ensures that the length of the file name is not too long.
	 * @param filename String
	 * @return String
	 */
	private static String checkFileNameLength(String filename) {
		// ensure that there is room the file suffix in length of the name
		final int maxlength = 31 - EXCEL_SUFFIX.length();
		
		if (filename.length() > maxlength) {
			// Truncate the name of the file
			filename = filename.substring(0, maxlength);
		}
		
		return filename;
	}
	
	/**
	 * Ensures that the file name ends with the correct suffix.
	 * @param filename String
	 * @return String 
	 */
	private static String checkFileNameSuffix(String filename) {
		
		if (!filename.endsWith(EXCEL_SUFFIX)) {
			filename = filename + EXCEL_SUFFIX;
		}
		
		return filename;
	}

	/**
	 * Tests that file location is 
	 * @param file File
	 * @throws EEHException
	 */
	protected static void testFileLocationWriteable(File file) throws EEHException {
		File location = new File(file.getParent());

		// if location is not writable 
		if (!location.canWrite() || !file.canWrite()) {
			throw new EEHException("File must be at a location with write permission.");
		}
	}
	
}
