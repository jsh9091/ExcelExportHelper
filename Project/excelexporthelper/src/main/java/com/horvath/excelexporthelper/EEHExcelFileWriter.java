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
import java.io.OutputStream;
import java.net.MalformedURLException;
import java.net.URISyntaxException;
import java.net.URL;
import java.nio.file.Files;
import java.util.List;

import org.apache.poi.common.usermodel.HyperlinkType;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Hyperlink;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 * Performs operations for preparing data for Excel file generation.
 * @author jhorvath
 */
final public class EEHExcelFileWriter {
	
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
		int rowNum = 0;
		
		// if we have a header row
		if (!eehSheet.getHeaders().isEmpty()) {
            Row row = xssfSheet.createRow(rowNum++);
            
            // setup bold style for use in header row cells
            CellStyle style = workbook.createCellStyle();
            Font font = workbook.createFont();
            font.setBold(true);
            style.setFont(font);
            
            int colNum = 0;
            // set the header row cells
            for (String header : eehSheet.getHeaders()) {
                Cell cell = row.createCell(colNum++);
                cell.setCellValue(header);
                cell.setCellStyle(style);
            }
		}
		
		// set cell data
		for (List<String> rowData : eehSheet.getData()) {
            Row row = xssfSheet.createRow(rowNum++);

            int colNum = 0;
            for (String data : rowData) {
                Cell cell = row.createCell(colNum++);
            	
                if (data == null) {
                	continue;
                	
            	} else if (canParseDouble(data)) {
        			double num = Double.parseDouble(data);
        			cell.setCellValue(num);
        			
            	} else if (canParseUrl(data)) {
            		// set the URL link data 
            		CreationHelper createHelper = workbook.getCreationHelper();
            		Hyperlink link = createHelper.createHyperlink(HyperlinkType.URL);
            		link.setAddress(data);
            		
            		// set style for URL link
            		CellStyle linkStyle = workbook.createCellStyle();
            		Font linkFont = workbook.createFont();
            		linkFont.setUnderline(Font.U_SINGLE);
            		linkFont.setColor(IndexedColors.BLUE.getIndex());
            		linkStyle.setFont(linkFont);
            		
            		// set the cell 
            		cell.setCellValue(data);
            		cell.setHyperlink(link);
            		cell.setCellStyle(linkStyle);
            		
            	} else if (canParseBoolean(data.trim())) {
            		boolean bool = Boolean.parseBoolean(data.trim());
            		cell.setCellValue(bool);
            		
            	} else {
            		// set data in the cell as a string
                    cell.setCellValue(data);
            	}
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
	
	/**
	 * Helper method to test if given string value 
	 * can be parsed to a numeric value or not.
	 * @param text String
	 * @return boolean 
	 */
	private boolean canParseDouble(String text) {
		boolean canParseDouble = true;
		
		try {
			Double.parseDouble(text);
		} catch (NumberFormatException ex) {
			canParseDouble = false;
		}
		return canParseDouble;
	}
	
	/**
	 * Helper method to test if given string value
	 * can be parsed to a URL address or not.
	 * @param text String 
	 * @return boolean 
	 */
	private boolean canParseUrl(String text) {
		boolean canParseUrl = true;

		try {
			URL url = new URL(text);
			url.toURI();
		} catch (URISyntaxException | MalformedURLException ex) {
			canParseUrl = false;
		}
		return canParseUrl;
	}
	
	/**
	 * Determines if the given string value can be 
	 * parsed to a boolean value or not. 
	 * @param text String
	 * @return boolean 
	 */
	private boolean canParseBoolean(String text) {
		boolean canParseBoolean = false;

		if ("true".equalsIgnoreCase(text) || "false".equalsIgnoreCase(text)) {
			canParseBoolean = true;
		}
		return canParseBoolean;
	}

}
