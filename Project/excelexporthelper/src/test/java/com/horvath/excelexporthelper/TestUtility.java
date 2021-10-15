/*
 * Joshua Horvath jsh_197@hotmail.com 
 */

package com.horvath.excelexporthelper;

import java.io.File;

import org.junit.Assert;

public class TestUtility {
	
	/**
	 * Creates folder and file with given names. 
	 * The folder is actually created at the user.dir location.
	 * @param testDirecory String 
	 * @param fileName String 
	 * @return File 
	 */
	public static File createValidFile(String testDirecory, String fileName) {
		String userdir = System.getProperty("user.dir");

		File folder = new File(userdir + File.separator + testDirecory);
		
		// if a folder from a failed run exists
		if (folder.exists()) {
			Assert.assertTrue(folder.delete());
		}
		
		Assert.assertTrue(folder.mkdir());
		
		return new File(folder + File.separator + fileName);
	}
	
	/**
	 * Deletes the parent folder of the given file. 
	 * @param file File 
	 */
	public static void cleanupParentFolder(File file) {
		Assert.assertTrue(file.getParentFile().delete());
	}

}
