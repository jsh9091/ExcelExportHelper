/*
 * Joshua Horvath jsh_197@hotmail.com 
 */

package com.horvath.excelexporthelper;

import java.util.ArrayList;

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
			Assert.assertTrue(sheet.getSheetName().length() == EEHSheet.MAX_NAME_LENGTH);
			
		} catch (EEHException e) {
			Assert.fail();
		}
	}
	
	@Test 
	public void EEHSheet_IllegalCharacters_Fixed() {
		// these characters are not allowed in an Excel sheet name
		final char[] illegalChars = { '\\', '/', '?', '*', '[', ']' };

		// NOTE: the illegal name does not already use the replacement character
		final String illegalName = "'This" + illegalChars[0] + illegalChars[1] + illegalChars[2] + illegalChars[3]
				+ illegalChars[4] + illegalChars[5] + "name is bad.'";
		
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
		
	}
	
	@Test
	public void EEHSheet_DuplicateNameX5_NameIncremented() { 
		
	}
}
