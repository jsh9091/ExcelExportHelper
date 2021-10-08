/*
 * Joshua Horvath jsh_197@hotmail.com 
 */

package com.horvath.excelexporthelper;

import java.io.File;

import org.junit.Assert;
import org.junit.Test;

public class FileUtilityTest {
	
	@Test
	public void validateFile_longname_nametruncated() throws EEHException {
	
		final String longFileName = "Thisfilenameislongerthanthirtyonecharactersintotallength.xlsx";
		
		final String actual = FileUtility.validateFileName(longFileName);

		Assert.assertEquals(31, actual.length());
	}
	
	@Test 
	public void validateFile_nosuffix_suffixadded() throws EEHException {
		final String noSuffix = "filename";
		final String actual = FileUtility.validateFileName(noSuffix);
		
		Assert.assertEquals(noSuffix + FileUtility.EXCEL_SUFFIX, actual); 
	}
	
	@Test
	public void validateFile_wrongsuffix_suffixadded() throws EEHException {
		final String wrongSuffix = "filename.pdf";
		final String actual = FileUtility.validateFileName(wrongSuffix);
		
		Assert.assertEquals(wrongSuffix + FileUtility.EXCEL_SUFFIX, actual); 
	}
	
	@Test 
	public void validateFile_emptyFileName_EEHException() {
		boolean catchException = false; 
		
		try {
			FileUtility.validateFileName("");
			Assert.fail(); // should not get here
		} catch (EEHException ex) {
			catchException = true;
		}
		Assert.assertTrue(catchException);
	}
	
	@Test
	public void ExcelExportHelper_emptyString_IllegalArgumentException() {
		// TODO: move this to another class
		boolean catchException = false; 
		
		try {
			new ExcelExportHelper("");
			Assert.fail(); // should not get here
		} catch (IllegalArgumentException actual) {
			catchException = true;
			Assert.assertEquals(FileUtility.EXCEPTION_EMPTY_STRING, actual.getMessage());
		}
		Assert.assertTrue(catchException);
	}
	
	@Test
	public void ExcelExportHelper_NonWritableLocation_IllegalArgumentException() {
		// TODO: move this to another class
		boolean catchException = false; 

		String userdir = System.getProperty("user.dir");
		File file = new File(userdir + File.separator + "TestDir_ExcelExportHelper");
		
		if (file.exists()) {
			// cleanup after a failed run
			file.delete();
		}
		
		Assert.assertTrue(file.mkdir());
		Assert.assertTrue(file.setReadOnly());
		

		try {
			new ExcelExportHelper(file.getAbsolutePath());
			Assert.fail(); // should not get here
		} catch (IllegalArgumentException actual) {
			catchException = true;
			Assert.assertEquals(FileUtility.EXCEPTION_WRITE_PERMISSION, actual.getMessage());
		}

		Assert.assertTrue(catchException);
		Assert.assertTrue(file.delete());
	}

	@Test
	public void testFileLocationWriteable_ValidLocation_Nothing() {
		String userdir = System.getProperty("user.dir");
		File file = new File(userdir);

		try {
			FileUtility.testFileLocationWriteable(file);
		} catch (EEHException e) {
			e.printStackTrace();
			Assert.fail(); // should not get here
		}
	}
	
	@Test
	public void testFileLocationWriteable_InvalidLocation_EEHException() {
		boolean catchException = false; 

		String userdir = System.getProperty("user.dir");
		File file = new File(userdir + File.separator + "TestDir_FileUtility");
		
		if (file.exists()) {
			// cleanup after a failed run
			file.delete();
		}
		
		Assert.assertTrue(file.mkdir());
		Assert.assertTrue(file.setReadOnly());

		try {
			FileUtility.testFileLocationWriteable(file);
			Assert.fail(); // should not get here
		} catch (EEHException actual) {
			catchException = true;
			Assert.assertEquals(FileUtility.EXCEPTION_WRITE_PERMISSION, actual.getMessage());
		}

		Assert.assertTrue(catchException);
		Assert.assertTrue(file.delete());
	}
}
