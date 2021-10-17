/*
 * Joshua Horvath jsh_197@hotmail.com 
 */

package com.horvath.excelexporthelper;

import java.io.File;
import java.util.ArrayList;

import org.junit.Assert;
import org.junit.Test;

/**
 * Unit tests for the main EEH library class. 
 * @author jhorvath
 */
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

		File file = TestUtility.createValidFile("TestDir_ExcelExportHelper", "BadLocation.xlsx");
		Assert.assertTrue(file.getParentFile().setReadOnly());
		
		try {
			new ExcelExportHelper(file.getAbsolutePath());
			Assert.fail(); // should not get here
		} catch (IllegalArgumentException actual) {
			caughtException = true;
			Assert.assertEquals(FileUtility.EXCEPTION_WRITE_PERMISSION, actual.getMessage());
		}

		Assert.assertTrue(caughtException);
		TestUtility.cleanupParentFolder(file);
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
	
	@Test
	public void createSheet_TooManySheets_IllegalArgumentException() {
		boolean caughtException = false; 
		File file = TestUtility.createValidFile("TooManySheets", "TestToomanySheets.xlsx");

		try {
			ExcelExportHelper eeh = new ExcelExportHelper(file.getAbsolutePath());
			
			for (int i = 0; i < EEHSheet.MAX_SHEET_COUNT + 2; i++) {
				eeh.createSheet("SheetA");
			}
			Assert.fail(); // should not get here
			
		} catch (IllegalArgumentException actual) {
			caughtException = true;
			Assert.assertEquals(EEHSheet.EXCEPTION_MAX_NUMBER_SHEETS_EXCEEDED, actual.getMessage());
		}
		Assert.assertTrue(caughtException);
		TestUtility.cleanupParentFolder(file);
	}
	
	@Test
	public void writeWorkBook_NoSheets_IllegalStateException() {
		boolean caughtException = false; 
		File file = TestUtility.createValidFile("NoSheets", "NoSheetsWriteTest.xlsx");

		try {
			ExcelExportHelper eeh = new ExcelExportHelper(file.getAbsolutePath());
			eeh.writeWorkBook();

			Assert.fail(); // should not get here
			
		} catch (IllegalStateException actual) {
			caughtException = true;
			Assert.assertEquals(ExcelExportHelper.EXCEPTION_NO_SHEETS_TO_WRITE, actual.getMessage());
		} catch (EEHException ex) {
			Assert.fail();
		}
		Assert.assertTrue(caughtException);
		TestUtility.cleanupParentFolder(file);
	}
	
	@Test
	public void writeWorkBook_VaildButEmptySheeets_FileWritten() {
		File file = TestUtility.createValidFile("NamedSheets", "NamedSheetsWriteTest.xlsx");

		try {
			ExcelExportHelper eeh = new ExcelExportHelper(file.getAbsolutePath());
			
			eeh.createSheet("SheetA");
			eeh.createSheet("SheetB");
			eeh.createSheet("Sheet*[]C'");
			
			eeh.writeWorkBook();
			
			Assert.assertTrue(file.exists());
			
			TestUtility.compareFileToData(eeh, file);

		} catch (EEHException ex) {
			Assert.fail();
		}
		TestUtility.cleanupParentFolder(file);
	}
	
	@Test
	public void writeWorkBook_PopulatedSheeets_FileWritten() {
		File file = TestUtility.createValidFile("PopulatedSheets", "PopSheetWriteTest.xlsx");

		try {
			ExcelExportHelper eeh = new ExcelExportHelper(file.getAbsolutePath());
			
			EEHSheet sheet = eeh.createSheet("Sheet A");
			ArrayList<String> data = new ArrayList<>();
			data.add("One");
			data.add("Two");
			data.add("Three");
			sheet.getData().add(data);

			data = new ArrayList<>();
			data.add("Four");
			data.add("Five");
			data.add("Six");
			data.add("Seven");
			sheet.getData().add(data);

			data = new ArrayList<>();
			data.add("Eight");
			sheet.getData().add(data);
			
			data = new ArrayList<>();
			data.add("Nine");
			data.add("Ten");
			data.add("Eleven");
			sheet.getData().add(data);

			eeh.writeWorkBook();
			
			Assert.assertTrue(file.exists());
			
			TestUtility.compareFileToData(eeh, file);

		} catch (EEHException ex) {
			Assert.fail();
		}
		TestUtility.cleanupParentFolder(file);
	}
}
