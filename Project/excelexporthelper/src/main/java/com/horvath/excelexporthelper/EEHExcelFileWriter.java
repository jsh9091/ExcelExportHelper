/*
 * Joshua Horvath jsh_197@hotmail.com 
 */

package com.horvath.excelexporthelper;

import java.io.File;
import java.io.IOException;
import java.io.OutputStream;
import java.nio.file.Files;
import java.util.List;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
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
		
		// set cell data
		int rowNum = 0;
		for (List<String> rowData : eehSheet.getData()) {
            Row row = xssfSheet.createRow(rowNum++);

            int colNum = 0;
            for (String data : rowData) {
                Cell cell = row.createCell(colNum++);
                cell.setCellValue(data);
            }
		}
		
	}
	
	/**
	 * Writes the Excel file to the disk.
	 * @param workbook XSSFWorkbook
	 * @throws EEHException
	 */
	private void generateFile(XSSFWorkbook workbook) throws EEHException {
		
        try (OutputStream os = Files.newOutputStream(this.file.toPath())) {
        	
            workbook.write(os);
            workbook.close();
            
        } catch (IOException ex) {
            throw new EEHException("Unexpected IO exception. " + ex.getMessage(), ex);
        }
	}

}
