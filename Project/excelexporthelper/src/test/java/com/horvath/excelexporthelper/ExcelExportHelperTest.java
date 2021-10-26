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
	
	@Test 
	public void writeWorkBook_HeaderRow_FileWritten() {
		File file = TestUtility.createValidFile("HeaderRowWrite", "HeaderRowWriteTest.xlsx");

		try {
			ExcelExportHelper eeh = new ExcelExportHelper(file.getAbsolutePath());
			
			EEHSheet sheet = eeh.createSheet("Header Sheet");
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
			
			sheet.getHeaders().add("Col 1");
			sheet.getHeaders().add("Col Two");
			sheet.getHeaders().add("Col 3.0");
			sheet.getHeaders().add("Four");
			sheet.getHeaders().add("5th Column");

			eeh.writeWorkBook();
			
			Assert.assertTrue(file.exists());
			
			TestUtility.compareFileToData(eeh, file);

		} catch (EEHException ex) {
			Assert.fail();
		}
		TestUtility.cleanupParentFolder(file);
	}

	
	@Test 
	public void writeWorkBook_NumberSetForCells_FileWritten() {
		File file = TestUtility.createValidFile("ParseNumber", "ParseNumberWriteTest.xlsx");

		try {
			ExcelExportHelper eeh = new ExcelExportHelper(file.getAbsolutePath());
			
			EEHSheet sheet = eeh.createSheet("Parse Num Sheet");
			ArrayList<String> data = new ArrayList<>();
			data.add("One");
			data.add("Two");
			data.add("Three");
			sheet.getData().add(data);

			data = new ArrayList<>();
			data.add("Numbers row:");
			data.add("5");
			data.add("6");
			data.add("7.0");
			data.add("7.1");
			sheet.getData().add(data);

			data = new ArrayList<>();
			data.add("Eight");
			sheet.getData().add(data);
			
			data = new ArrayList<>();
			data.add("Nine 9");
			data.add("Ten 10.0");
			data.add("11");
			data.add("tizenketto");
			sheet.getData().add(data);

			eeh.writeWorkBook();
			
			Assert.assertTrue(file.exists());
			
			TestUtility.compareFileToData(eeh, file);

		} catch (EEHException ex) {
			Assert.fail();
		}
		TestUtility.cleanupParentFolder(file);
	}
	
	@Test 
	public void writeWorkBook_UrlSetForCells_FileWritten() {
		File file = TestUtility.createValidFile("ParseURL", "ParseURLWriteTest.xlsx");

		try {
			ExcelExportHelper eeh = new ExcelExportHelper(file.getAbsolutePath());
			
			EEHSheet sheet = eeh.createSheet("Parse URL Sheet");
			ArrayList<String> data = new ArrayList<>();
			data.add("One");
			data.add("Two");
			data.add("Three");
			sheet.getData().add(data);

			data = new ArrayList<>();
			data.add("https://poi.apache.org/");
			data.add("https://www.google.com/");
			data.add("https://slashdot.org/");
			sheet.getData().add(data);

			eeh.writeWorkBook();
			
			Assert.assertTrue(file.exists());
			
			TestUtility.compareFileToData(eeh, file);

		} catch (EEHException ex) {
			Assert.fail();
		}
		TestUtility.cleanupParentFolder(file);
	}
	
	@Test 
	public void writeWorkBook_BooleanSetForCells_FileWritten() {
		File file = TestUtility.createValidFile("ParseBoolean", "ParseBooleanWriteTest.xlsx");

		try {
			ExcelExportHelper eeh = new ExcelExportHelper(file.getAbsolutePath());
			
			EEHSheet sheet = eeh.createSheet("Parse Boolean Sheet");
			ArrayList<String> data = new ArrayList<>();
			data.add("One");
			data.add("Two");
			data.add("Three");
			sheet.getData().add(data);

			data = new ArrayList<>();
			data.add("True"); // true
			data.add("TRUE"); // true
			data.add("true"); // true
			data.add("False"); // false
			data.add("FALSE"); // false
			data.add("False"); // false
			data.add("  True"); // false
			data.add("TRUE  "); // false
			data.add("  true  "); // false
			data.add("1 True"); // string
			data.add("TRUE  2 "); // string
			data.add("1  true  a"); // string
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
