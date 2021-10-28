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
import java.io.IOException;
import java.io.InputStream;
import java.nio.file.Files;
import java.util.List;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
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
		} catch (IOException ex) {
            System.err.println(ex.getMessage());
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
		
		try (InputStream is = Files.newInputStream(file.toPath())) {
			
			 workbook = new XSSFWorkbook(is);
			
		} catch (IOException ex) {
            System.err.println(ex.getMessage());
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

		// check the name of the sheets
		Assert.assertEquals(eehSheet.getSheetName(), xssfSheet.getSheetName());

		// check the row counts
		if (eehSheet.getHeaders().isEmpty()) {
			Assert.assertEquals(eehSheet.getData().size(), calculateRowCount(xssfSheet));
		} else {
			Assert.assertEquals(eehSheet.getData().size() + 1, calculateRowCount(xssfSheet));
		}

		int rowNum = 0;

		// if our sheet has header cells
		if (!eehSheet.getHeaders().isEmpty()) {
            Row row = xssfSheet.getRow(rowNum++);
            int colNum = 0;

            for (String header : eehSheet.getHeaders()) {
                Cell cell = row.getCell(colNum++);
                
                // compare the header cell values
                Assert.assertEquals(header, cell.getStringCellValue());
			}
		}

		// compare the sheet data
		for (List<String> rowData : eehSheet.getData()) {
            Row row = xssfSheet.getRow(rowNum++);

            int colNum = 0;
            for (String data : rowData) {
                Cell cell = row.getCell(colNum++);
                
                // compare actual vs control based on actual cell type
                switch (cell.getCellType()) {
                
                case NUMERIC:
                	double control = Double.parseDouble(data);
                	double actual = cell.getNumericCellValue();
                	Assert.assertEquals(control, actual, 0.0001);
                	break;
                	
                case STRING:
                    // compare the cell values
                    Assert.assertEquals(data, cell.getStringCellValue());
                	break;
                	
                case BOOLEAN:
                	boolean bool = Boolean.parseBoolean(data.trim());
                	Assert.assertEquals(bool, cell.getBooleanCellValue());
                	break;
                	
                case _NONE:
                case BLANK:
                case ERROR:

                default:
                	break;
                }
            }
		}
	}
	
	/**
	 * Calculate the number of rows in the given XSSFSheet. 
	 * @param xssfSheet XSSFSheet
	 * @return int
	 */
	private static int calculateRowCount(XSSFSheet xssfSheet) {
		int rowTotal = xssfSheet.getLastRowNum();

		if ((rowTotal > 0) || (xssfSheet.getPhysicalNumberOfRows() > 0)) {
		    rowTotal++;
		} else if (rowTotal < 0) {
			rowTotal = 0;
		}
		return rowTotal;
	}
}
