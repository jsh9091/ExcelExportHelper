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

import java.util.ArrayList;
import java.util.List;
import org.junit.Assert;
import org.junit.Test;

/**
 * Performs tests on sheet class. 
 * @author jhorvath
 */
public class EEHSheetTest {

	@Test
	public void EEHSheet_ValidName_SheetCreated() {
		final String sheetName = "Sheet A";
		try {
			EEHSheet sheet = new EEHSheet(sheetName, new ArrayList<EEHSheet>());
			Assert.assertNotNull(sheet);
			Assert.assertEquals(sheetName, sheet.getSheetName());
			
		} catch (EEHException e) {
			Assert.fail();
		}
	}
	
	@Test
	public void EEHSheet_NullString_EEHException() {
		boolean caughtException = false;
		try {
			// pass null as sheet name 
			new EEHSheet(null, new ArrayList<EEHSheet>());
			Assert.fail();
		} catch (EEHException actual) {
			caughtException = true;
			Assert.assertEquals(EEHSheet.EXCEPTION_EMPTY_OR_NULL_SHEETNAME, actual.getMessage());
		}
		Assert.assertTrue(caughtException);
	}
	
	@Test
	public void EEHSheet_EmptyString_EEHException() {
		boolean caughtException = false;
		try {
			// empty string as sheet name
			new EEHSheet("", new ArrayList<EEHSheet>());
			Assert.fail();
		} catch (EEHException actual) {
			caughtException = true;
			Assert.assertEquals(EEHSheet.EXCEPTION_EMPTY_OR_NULL_SHEETNAME, actual.getMessage());
		}
		Assert.assertTrue(caughtException);
	}
	
	@Test
	public void EEHSheet_LongName_Fixed() {
		final String sheetName = "Sheet Long Name 1234567890123456789012345678901234567890";
		try {
			EEHSheet sheet = new EEHSheet(sheetName, new ArrayList<EEHSheet>());
			
			Assert.assertNotNull(sheet);
			Assert.assertNotEquals(sheetName, sheet.getSheetName());
			// the name has been truncated to the max sized name for a sheet
			Assert.assertEquals(EEHSheet.MAX_NAME_LENGTH, sheet.getSheetName().length());
			
		} catch (EEHException e) {
			Assert.fail();
		}
	}
	
	@Test 
	public void EEHSheet_IllegalCharacters_Fixed() {
		// these characters are not allowed in an Excel sheet name
		final char[] illegalChars = { '\\', '/', '?', '*', '[', ']', ':' };
		
		StringBuilder sb = new StringBuilder();
		sb.append("'"); // illegal character at start of name
		sb.append("This");
		for (int i = 0; i < illegalChars.length; i++) {
			sb.append(illegalChars[i]);
		}
		sb.append("name is bad.");
		sb.append("'"); // illegal character at end of name
		
		// NOTE: the illegal name does not already use the replacement character
		final String illegalName = sb.toString();
		
		try {
			EEHSheet sheet = new EEHSheet(illegalName, new ArrayList<EEHSheet>());
			
			Assert.assertNotNull(sheet);
			// the sheet name has been altered
			Assert.assertNotEquals(illegalName, sheet.getSheetName());
			// the length of the sheet has not changed
			Assert.assertEquals(illegalName.length(), sheet.getSheetName().length());
			// get the number of times the replacement character is now in the name
			final int replacementCount = countCharsInString(sheet.getSheetName(), '-');
			// + 2 on array length to account for starting and ending apostrophes in name
			Assert.assertEquals(illegalChars.length + 2, replacementCount);
			
		} catch (EEHException e) {
			Assert.fail();
		}
	}
	
	/**
	 * Helper method that counts the occurrence of a character in a given string. 
	 * @param text String 
	 * @param character char 
	 * @return int 
	 */
	private int countCharsInString(String text, char character) {
		int count = 0;
		 
		for (int i = 0; i < text.length(); i++) {
		    if (text.charAt(i) == character) {
		        count++;
		    }
		}
		return count;
	}
	
	@Test
	public void EEHSheet_DuplicateName_NameIncremented() { 
		
		final String sheetName = "Duplicate";
		try {
			List<EEHSheet> sheets = new ArrayList<>(2);
			
			EEHSheet sheet = new EEHSheet(sheetName, sheets);
			sheets.add(sheet);
			
			EEHSheet sheet2 = new EEHSheet(sheetName, sheets);
			sheets.add(sheet2);
			
			Assert.assertEquals(2, sheets.size());
			
			Assert.assertEquals(sheetName, sheet.getSheetName());
			Assert.assertEquals(sheetName + "1", sheet2.getSheetName());
			
		} catch (EEHException e) {
			Assert.fail();
		}
	}
	
	@Test
	public void EEHSheet_DuplicateNameX5_NameIncremented() { 
		
		final String sheetName = "Dup";
		final int limit = 5;
		
		try {
			List<EEHSheet> sheets = new ArrayList<>(5);
			
			// create five sheets with the same name
			for (int i = 0; i < limit; i++) {
				EEHSheet sheet = new EEHSheet(sheetName, sheets);
				sheets.add(sheet);
			}
						
			Assert.assertEquals(limit, sheets.size());
						
			// verify that the names of the sheets after the first one have incremented names
			for (int i = 0; i < sheets.size(); i++) {
				if (i == 0) {
					Assert.assertEquals(sheetName, sheets.get(i).getSheetName());
				} else {
					Assert.assertEquals(sheetName + i, sheets.get(i).getSheetName());
				}
			}
			
		} catch (EEHException e) {
			Assert.fail();
		}
	}
	
	@Test
	public void EEHSheet_DuplicateNameTooManySheets_NameIncremented() { 

		boolean caughtException = false;
		final String sheetName = "MaxDup";
		final int controlLimit = 2;
		
		try {
			List<EEHSheet> sheets = new ArrayList<>(EEHSheet.MAX_SHEET_COUNT + controlLimit);
			
			// create many sheets with the same name
			for (int i = 0; i < EEHSheet.MAX_SHEET_COUNT + controlLimit; i++) {
				EEHSheet sheet = new EEHSheet(sheetName, sheets);
				sheets.add(sheet);
			}
			Assert.fail(); // should not get here

		} catch (EEHException actual) {
			caughtException = true;
			Assert.assertEquals(EEHSheet.EXCEPTION_MAX_NUMBER_SHEETS_EXCEEDED, actual.getMessage());
		}
		Assert.assertTrue(caughtException);
	}
	
}
