/*
 * Joshua Horvath jsh_197@hotmail.com 
 */

package com.horvath.excelexporthelper;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.List;

import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 * Performs operations for preparing data for Excel file generation.
 * @author jhorvath
 */
public class EEHExcelFileWriter {
	
	private File file;
	private List<EEHSheet> sheets;
	
	/**
	 * Constructor. 
	 * @param file File
	 * @param sheets List<EEHSheet>
	 */
	protected EEHExcelFileWriter(File file, List<EEHSheet> sheets) {
		this.file = file;
		this.sheets = sheets;
	}
	
	/**
	 * Triggers operations to prepare data for Excel 
	 * file generation and writes the file.
	 * @throws EEHException 
	 */
	protected void writeFile() throws EEHException {
		
        XSSFWorkbook workbook = new XSSFWorkbook();
        
        for (EEHSheet eehSheet : this.sheets) {
        	// create and populate XSSFSheets 
        	createSheet(eehSheet, workbook);
        }
        
        // TODO sheet data method needs to be forward thinking to be expanded 
        // for different data types when putting in cell data is a thing
        
        // perform file writing operations 
        generateFile(workbook);
	}
	
	/**
	 * Creates the Excel sheet and populates it data.
	 * @param eehSheet EEHSheet
	 * @param workbook XSSFWorkbook
	 */
	private void createSheet(EEHSheet eehSheet, XSSFWorkbook workbook) {
		
		XSSFSheet xssfSheet = workbook.createSheet(eehSheet.getSheetName());
		
		// TODO: set cell data
		
	}
	
	/**
	 * Writes the Excel file to the disk.
	 * @param workbook XSSFWorkbook
	 * @throws EEHException
	 */
	private void generateFile(XSSFWorkbook workbook) throws EEHException {
		
        try (FileOutputStream outputStream = new FileOutputStream(this.file)) {
        	
            workbook.write(outputStream);
            workbook.close();
            
        } catch (IOException ex) {
            ex.printStackTrace();
            throw new EEHException("Unexpected IO exception. " + ex.getMessage());
        }
	}

}
