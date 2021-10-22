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

import java.util.ArrayList;
import java.util.List;

import org.apache.poi.ss.util.WorkbookUtil;

/**
 * Defines the data for an individual Excel sheet. 
 * @author jhorvath
 */
final public class EEHSheet {
	
	private String sheetName;
	private List<ArrayList<String>> data;
	
	public static String EXCEPTION_EMPTY_OR_NULL_SHEETNAME = "Sheet name not be null or empty.";
	public static String EXCEPTION_MAX_NUMBER_SHEETS_EXCEEDED = "The maximum number of sheets in an Excel file has been exceeded.";
	
	/**
	 * The maximum number of sheets allowed in an Excel file.
	 */
	public static int MAX_SHEET_COUNT = 255;
	
	/**
	 * The maximum number of characters allowed in a sheet name.
	 */
	public static int MAX_NAME_LENGTH = 31;
	
	/**
	 * Constructor for a sheet. 
	 * @param sheetName String Name for the sheet. 
	 * @param sheets List<EEHSheet> Existing sheets in current instance. 
	 * @throws EEHException
	 */
	protected EEHSheet(String sheetName, List<EEHSheet> sheets) throws EEHException {
		
		if (sheets.size() >= MAX_SHEET_COUNT) {
			throw new EEHException(EXCEPTION_MAX_NUMBER_SHEETS_EXCEEDED);
		}
		
		this.sheetName = createSafeSheetName(sheetName, sheets);
		this.data = new ArrayList<ArrayList<String>>();
	}
	
	/**
	 * As needed performs adjustments to given sheet name to 
	 * avoid illegal characters or duplicate names. 
	 * @param sheetName String 
	 * @param sheets List<EEHSheet>
	 * @return String 
	 * @throws EEHException
	 */
	private String createSafeSheetName(String sheetName, List<EEHSheet> sheets) throws EEHException {
		
		if (sheetName == null || sheetName.isEmpty()) {
			throw new EEHException(EXCEPTION_EMPTY_OR_NULL_SHEETNAME);
		}

		// replace illegal characters 
		String result = WorkbookUtil.createSafeSheetName(sheetName, '-');
		
		if (sheets != null && !sheets.isEmpty()) {
			// we need a list of the current sheet names to avoid duplicate names
			List<String> currentSheetNames = new ArrayList<>(sheets.size());
			
			// populate the list of current sheet names
			for (EEHSheet sheet : sheets) {
				currentSheetNames.add(sheet.getSheetName());
			}
			
			// if the given name is a duplicate, then fix it
			result = fixDuplicateName(result, currentSheetNames, 1);
		}
		
		return result;
	}
	
	/**
	 * If the given name is a name that already exists, 
	 * then append the given integer to end of name.
	 * @param sheetName String 
	 * @param currentSheetNames List<String> 
	 * @param count int 
	 * @return String 
	 */
	private String fixDuplicateName(String sheetName, List<String> currentSheetNames, int count) {
		// update the name
		String newName = sheetName + count;
		
		// verify that the updated name doesn't already exist in the instance
		if (currentSheetNames.contains(newName)) {
			if (count < MAX_SHEET_COUNT) {
				// recurse 
				return fixDuplicateName(sheetName, currentSheetNames, ++count);
			}
		}
			
		return newName;
	}
	
	public String getSheetName() {
		return this.sheetName;
	}

	public List<ArrayList<String>> getData() {
		return data;
	}

	@Override
	public String toString() {
		return "EEHSheet [sheetName=" + sheetName + ", data=" + data + "]";
	}

}
