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
		boolean caughtException = false; 
		
		try {
			new ExcelExportHelper("");
			Assert.fail(); // should not get here
		} catch (IllegalArgumentException actual) {
			caughtException = true;
			Assert.assertEquals(FileUtility.EXCEPTION_EMPTY_STRING, actual.getMessage());
		}
		Assert.assertTrue(caughtException);
	}
	
	@Test
	public void ExcelExportHelper_NonWritableLocation_IllegalArgumentException() {
		boolean caughtException = false; 

		String userdir = System.getProperty("user.dir");
		File folder = new File(userdir + File.separator + "TestDir_ExcelExportHelper");
		
		if (folder.exists()) {
			// cleanup after a failed run
			folder.delete();
		}
		
		Assert.assertTrue(folder.mkdir());
		Assert.assertTrue(folder.setReadOnly());
		
		File file = new File(folder.getAbsolutePath() + File.separator + "Test.xlsx");

		try {
			new ExcelExportHelper(file.getAbsolutePath());
			Assert.fail(); // should not get here
		} catch (IllegalArgumentException actual) {
			caughtException = true;
			Assert.assertEquals(FileUtility.EXCEPTION_WRITE_PERMISSION, actual.getMessage());
		}

		Assert.assertTrue(caughtException);
		Assert.assertTrue(folder.delete());
	}
	
	@Test
	public void ExcelExportHelper_ValidLocation_Initialized() {

		File file = TestUtility.createValidFile("ValidLocationTest", "Test.xlsx");
		
		try {
			ExcelExportHelper actual = new ExcelExportHelper(file.getAbsolutePath());
			Assert.assertNotNull(actual);
		} catch (IllegalArgumentException e) {
			Assert.fail();
		}
		TestUtility.cleanupParentFolder(file);
	}

	@Test
	public void createSheet_EmptyString_IllegalArgumentException() {
		boolean caughtException = false;
		File file = TestUtility.createValidFile("sheetEmptyString", "Test.xlsx");
		
		try {
			ExcelExportHelper eeh = new ExcelExportHelper(file.getAbsolutePath());
			eeh.createSheet(""); // test
			Assert.fail(); // should not get here
		} catch (IllegalArgumentException actual) {
			caughtException = true;
			Assert.assertEquals(EEHSheet.EXCEPTION_EMPTY_OR_NULL_SHEETNAME, actual.getMessage());
		}
		Assert.assertTrue(caughtException);
		TestUtility.cleanupParentFolder(file);
	}

	@Test
	public void createSheet_NullString_IllegalArgumentException() {
		boolean caughtException = false;
		File file = TestUtility.createValidFile("nullEmptyString", "Test.xlsx");
		
		try {
			ExcelExportHelper eeh = new ExcelExportHelper(file.getAbsolutePath());
			eeh.createSheet(null); // test
			Assert.fail(); // should not get here
		} catch (IllegalArgumentException actual) {
			caughtException = true;
			Assert.assertEquals(EEHSheet.EXCEPTION_EMPTY_OR_NULL_SHEETNAME, actual.getMessage());
		}
		Assert.assertTrue(caughtException);
		TestUtility.cleanupParentFolder(file);
	}
	
	@Test
	public void createSheet_ValidName_SheetReturned() {
		File file = TestUtility.createValidFile("nullEmptyString", "Test.xlsx");
		
		try {
			ExcelExportHelper eeh = new ExcelExportHelper(file.getAbsolutePath());
			EEHSheet sheet = eeh.createSheet("New Sheet");
			Assert.assertNotNull(sheet);
		} catch (IllegalArgumentException actual) {
			Assert.fail(); // should not get here
		}
		TestUtility.cleanupParentFolder(file);
	}
	
	@Test
	public void createSheet_MultipeSheets_SheetAdded() {
		File file = TestUtility.createValidFile("nullEmptyString", "Test.xlsx");
		
		try {
			ExcelExportHelper eeh = new ExcelExportHelper(file.getAbsolutePath());
			
			EEHSheet sheet = eeh.createSheet("Sheet1");
			Assert.assertNotNull(sheet);

			sheet = eeh.createSheet("Sheet2");
			Assert.assertNotNull(sheet);
			
			sheet = eeh.createSheet("Sheet3");
			Assert.assertNotNull(sheet);

			sheet = eeh.createSheet("Sheet4");
			Assert.assertNotNull(sheet);
		} catch (IllegalArgumentException actual) {
			Assert.fail(); // should not get here
		}
		TestUtility.cleanupParentFolder(file);
	}
}
