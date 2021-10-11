/*
 * Joshua Horvath jsh_197@hotmail.com 
 */

package com.horvath.excelexporthelper;

import java.io.File;

import org.junit.Assert;
import org.junit.Test;

public class ExcelExportHelperTest {
	
	@Test
	public void ExcelExportHelper_emptyString_IllegalArgumentException() {
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
	public void createSheet_EmptyString_IllegalArgumentException() {
		
	}

	@Test
	public void createSheet_NullString_IllegalArgumentException() {
		
	}
	
	@Test
	public void createSheet_ValidName_SheetReturned() {
		
	}
	
	@Test
	public void createSheet_MultipeSheets_SheetAdded() {
		
	}
}
