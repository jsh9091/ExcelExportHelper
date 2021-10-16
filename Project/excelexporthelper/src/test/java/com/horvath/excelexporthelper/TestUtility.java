/*
 * Joshua Horvath jsh_197@hotmail.com 
 */

package com.horvath.excelexporthelper;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.Assert;

/**
 * Utility class for EEH unit test operations. 
 * @author jhorvath
 */
public class TestUtility {
	
	private TestUtility() {}
	
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
			Assert.assertTrue(deleteDirectory(folder));
		}
		
		Assert.assertTrue(folder.mkdir());
		
		return new File(folder + File.separator + fileName);
	}
	
	/**
	 * Deletes the parent folder of the given file. 
	 * @param file File 
	 */
	public static void cleanupParentFolder(File file) {
		Assert.assertTrue(deleteDirectory(file.getParentFile()));
	}
	
	/**
	 * Deletes the folder and all contents.
	 * @param folder File 
	 * @return boolean
	 */
	private static boolean deleteDirectory(File folder) {
	    File[] contents = folder.listFiles();
	    if (contents != null) {
	        for (File file : contents) {
	            deleteDirectory(file);
	        }
	    }
	    return folder.delete();
	}
	
	/**
	 * Compares the data in the written file to the control data
	 * in the EEH object reference. 
	 * @param eeh ExcelExportHelper
	 * @param file File 
	 */
	public static void compareFileToData(ExcelExportHelper eeh, File file) {
		
		Assert.assertEquals(eeh.getFile().getName(), file.getName());
		
		XSSFWorkbook workBook = readWorkbook(file);
		
		Assert.assertNotNull(workBook);
		
		Assert.assertEquals(eeh.getSheets().size(), workBook.getNumberOfSheets());
		
		// cycle over the sheets and start actual comparisons 
		for (int i = 0; i < workBook.getNumberOfSheets(); i++) {
			XSSFSheet xssfSheet = workBook.getSheetAt(i);
			
			compareSheetData(eeh.getSheets().get(i), xssfSheet);
		}
		
		try {
			workBook.close();
		} catch (IOException e) {
			e.printStackTrace();
			Assert.fail();
		}
	}

	/**
	 * Reads in the file from disk and returns an excel workbook.
	 * @param file File 
	 * @return XSSFWorkbook
	 */
	private static XSSFWorkbook readWorkbook(File file) {
		XSSFWorkbook workbook = null;
		Assert.assertTrue(file.exists());
		
		try (FileInputStream fis = new FileInputStream(file)) {
			
			 workbook = new XSSFWorkbook(fis);
			
		} catch (IOException e) {
            e.printStackTrace();
            Assert.fail();
        } 
		return workbook;
	}
	
	/**
	 * Compares the data for a given sheet. 
	 * The EEHSheet is treated as the control sheet, 
	 * and the XSSFSheet is treated as the actual 
	 * value to be tested.
	 * @param eehSheet EEHSheet 
	 * @param xssfSheet XSSFSheet
	 */
	private static void compareSheetData(EEHSheet eehSheet, XSSFSheet xssfSheet) {

		Assert.assertEquals(eehSheet.getSheetName(), xssfSheet.getSheetName());
		
	}
}
