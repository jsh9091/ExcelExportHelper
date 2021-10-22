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

import org.junit.Assert;
import org.junit.Test;

public class FileUtilityTest {
	
	@Test
	public void validateFile_longname_nametruncated() throws EEHException {
	
		final String longFileName = "Thisfilenameislongerthanthirtyonecharactersintotallength.xlsx";
		
		final String actual = FileUtility.validateFileName(longFileName);

		Assert.assertEquals(FileUtility.FILENAME_MAX_LENGTH, actual.length());
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
	public void testFileLocationWriteable_ValidLocation_Nothing() {
		String userdir = System.getProperty("user.dir");
		File file = new File(userdir);

		try {
			FileUtility.testFileLocationWriteable(file);
		} catch (EEHException ex) {
			System.err.println(ex.getMessage());
			Assert.fail(); // should not get here
		}
	}
	
	@Test
	public void testFileLocationWriteable_InvalidLocation_EEHException() {
		boolean catchException = false; 

		String userdir = System.getProperty("user.dir");
		File folder = new File(userdir + File.separator + "TestDir_FileUtility");
		
		if (folder.exists()) {
			// cleanup after a failed run
			Assert.assertTrue(folder.delete());
		}
		
		Assert.assertTrue(folder.mkdir());
		Assert.assertTrue(folder.setReadOnly());

		File file = new File(folder.getAbsolutePath() + File.separator + "Test.xlsx");
		
		try {
			FileUtility.testFileLocationWriteable(file);
			Assert.fail(); // should not get here
		} catch (EEHException actual) {
			catchException = true;
			Assert.assertEquals(FileUtility.EXCEPTION_WRITE_PERMISSION, actual.getMessage());
		}

		Assert.assertTrue(catchException);
		Assert.assertTrue(folder.delete());
	}
}
