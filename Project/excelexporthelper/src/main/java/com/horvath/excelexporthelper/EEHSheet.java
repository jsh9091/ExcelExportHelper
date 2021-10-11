/*
 * Joshua Horvath jsh_197@hotmail.com 
 */

package com.horvath.excelexporthelper;

import java.util.ArrayList;
import java.util.List;

import org.apache.poi.ss.util.WorkbookUtil;

final class EEHSheet {
	
	private String sheetName;
	
	public static String EXCEPTION_EMPTY_OR_NULL_SHEETNAME = "Sheet name not be null or empty.";
	
	
	protected EEHSheet(String sheetName, List<EEHSheet> sheets) throws EEHException {
		this.sheetName = createSafeSheetName(sheetName, sheets);
		
	}
	
	
	private String createSafeSheetName(String sheetName, List<EEHSheet> sheets) throws EEHException {
		
		if (sheetName == null || sheetName.isEmpty()) {
			throw new EEHException(EXCEPTION_EMPTY_OR_NULL_SHEETNAME);
		}

		String result = WorkbookUtil.createSafeSheetName(sheetName, '-');
		
		if (sheets != null && !sheets.isEmpty()) {
			List<String> currentSheetNames = new ArrayList<>(sheets.size());
			
			for (EEHSheet sheet : sheets) {
				currentSheetNames.add(sheet.getSheetName());
			}
			
			result = fixDuplicateName(sheetName, currentSheetNames, 1);
		}
		
		return result;
	}
	
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
