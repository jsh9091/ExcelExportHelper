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
	public void EEHSheet_ValidName_SheetReturned() {
		
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
		
	}
	
	@Test 
	public void EEHSheet_IllegalCharacters_Fixed() {
		
	}
	
	@Test
	public void EEHSheet_DuplicateName_NameIncremented() { 
		
	}
	
	@Test
	public void EEHSheet_DuplicateNameX5_NameIncremented() { 
		
	}
}
