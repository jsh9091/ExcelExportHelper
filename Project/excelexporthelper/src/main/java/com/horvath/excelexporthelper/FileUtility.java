/*
 * Joshua Horvath jsh_197@hotmail.com 
 */

package com.horvath.excelexporthelper;

import java.io.File;

/**
 * Utility class for file operations. 
 * @author jhorvath
 */
final class FileUtility {
	
	private FileUtility() {}
	
	protected static final String EXCEL_SUFFIX = ".xlsx";
	protected static final String EXCEPTION_WRITE_PERMISSION = "File must be at a location with write permission.";
	protected static final String EXCEPTION_EMPTY_STRING = "File name must not be empty.";
	protected static final int FILENAME_MAX_LENGTH = 31;
	
	/**
	 * Ensures that the name of the file is valid.
	 * @param filename String 
	 * @return String 
	 * @throws EEHException 
	 */
	protected static String validateFileName(String filename) throws EEHException {
		
		if (filename.isEmpty()) {
			throw new EEHException(EXCEPTION_EMPTY_STRING);
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
		final int maxlength = FILENAME_MAX_LENGTH - EXCEL_SUFFIX.length();
		
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
		if (!location.canWrite()) {
			throw new EEHException(EXCEPTION_WRITE_PERMISSION);
		}
	}
	
}
