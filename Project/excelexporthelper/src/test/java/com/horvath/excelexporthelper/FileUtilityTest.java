/*
 * Joshua Horvath jsh_197@hotmail.com 
 */

package com.horvath.excelexporthelper;

import static org.junit.Assert.assertEquals;
import static org.junit.Assert.assertTrue;
import static org.junit.Assert.fail;

import org.junit.Test;

public class FileUtilityTest {
	
	@Test
	public void validateFile_longname_nametruncated() throws EEHException {
	
		final String longFileName = "Thisfilenameislongerthanthirtyonecharactersintotallength.xlsx";
		
		final String actual = FileUtility.validateFileName(longFileName);

		assertEquals(31, actual.length());
	}
	
	@Test 
	public void validateFile_nosuffix_suffixadded() throws EEHException {
		final String noSuffix = "filename";
		final String actual = FileUtility.validateFileName(noSuffix);
		
		assertEquals(noSuffix + FileUtility.EXCEL_SUFFIX, actual); 
	}
	
	@Test
	public void validateFile_wrongsuffix_suffixadded() throws EEHException {
		final String wrongSuffix = "filename.pdf";
		final String actual = FileUtility.validateFileName(wrongSuffix);
		
		assertEquals(wrongSuffix + FileUtility.EXCEL_SUFFIX, actual); 
	}
	
	@Test 
	public void validateFile_emptyFileName_EEHExceptionThrown() {
		boolean catchException = false; 
		
		try {
			FileUtility.validateFileName("");
			fail(); // should not get here
		} catch (EEHException ex) {
			catchException = true;
		}
		assertTrue(catchException);
	}
	
	@Test
	public void ExcelExportHelper_emptyString_IllegalArgumentException() {
		boolean catchException = false; 
		
		try {
			new ExcelExportHelper("");
			fail(); // should not get here
		} catch (IllegalArgumentException ex) {
			catchException = true;
		}
		assertTrue(catchException);
	}
}
