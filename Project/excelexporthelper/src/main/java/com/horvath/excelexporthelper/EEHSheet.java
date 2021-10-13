/*
 * Joshua Horvath jsh_197@hotmail.com 
 */

package com.horvath.excelexporthelper;

import java.util.ArrayList;
import java.util.List;

import org.apache.poi.ss.util.WorkbookUtil;

/**
 * Defines the data for an individual Excel sheet. 
 * @author jhorvath
 */
final class EEHSheet {
	
	private String sheetName;
	
	public static String EXCEPTION_EMPTY_OR_NULL_SHEETNAME = "Sheet name not be null or empty.";
	
	/**
	 * Constructor for a sheet. 
	 * @param sheetName String Name for the sheet. 
	 * @param sheets List<EEHSheet> Existing sheets in current instance. 
	 * @throws EEHException
	 */
	protected EEHSheet(String sheetName, List<EEHSheet> sheets) throws EEHException {
		this.sheetName = createSafeSheetName(sheetName, sheets);
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
			result = fixDuplicateName(sheetName, currentSheetNames, 1);
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
		
		if (currentSheetNames.contains(sheetName)) {
			
			return fixDuplicateName(sheetName + count, currentSheetNames, ++count);
		}
			
		return sheetName;
	}
	
	
	public String getSheetName() {
		return this.sheetName;
	}

}
