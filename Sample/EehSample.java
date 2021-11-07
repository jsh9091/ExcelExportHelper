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

package com.horvath.eeh.sample;

import java.io.File;
import java.util.ArrayList;
import java.util.List;

import com.horvath.excelexporthelper.EEHException;
import com.horvath.excelexporthelper.EEHSheet;
import com.horvath.excelexporthelper.ExcelExportHelper;

/**
 * An example class demonstrating the usage of the Excel Export Helper tool. 
 * @author jhorvath
 */
public class EehSample {
	
	private static List<RadioStation> stations = new ArrayList<>();

	public static void main(String[] args) {
		
		// build up data
		buildRadioStationData();
		
		// create the file at the root level of the user folder
		File file = new File(System.getProperty("user.home") + "/EEH Sample File.xlsx");

		// create our EEH object reference 
		ExcelExportHelper eeh = new ExcelExportHelper(file);
		// create a sheet to populate 
		EEHSheet sheet = eeh.createSheet("Radio Stations");
		
		// set the headers for the sheet
		sheet.getHeaders().add("Call Letters");
		sheet.getHeaders().add("Frequency");
		sheet.getHeaders().add("Modulation");
		sheet.getHeaders().add("Format");
		sheet.getHeaders().add("Network Or Organization");
		sheet.getHeaders().add("Website");
				
		for (RadioStation station : stations) {
			// build data for a station into a list of Strings
			ArrayList<String> stationData = new ArrayList<>();
			stationData.add(station.callLetters);
			stationData.add(station.frequency);
			stationData.add(station.modulation);
			stationData.add(station.format);
			stationData.add(station.networkOrOrganization);
			stationData.add(station.website);
			// set the sheet data for a row
			sheet.getData().add(stationData);
		}
		
		try {
			// call command to write out the Excel file
			eeh.writeWorkBook();
			System.out.println("File written to: " + eeh.getFile().getAbsolutePath());
			
		} catch (EEHException ex) {
			System.err.println(ex.getMessage());
		}
	}
	
	/**
	 * Populates the internal list of ration stations.
	 */
	public static void buildRadioStationData() {
		
		RadioStation station = new RadioStation();
		station.callLetters = "WYSO";
		station.frequency = "91.3";
		station.modulation = "FM";
		station.format = "News";
		station.networkOrOrganization = "NPR";
		station.website = "https://www.wyso.org/";
		stations.add(station);
		
		station = new RadioStation();
		station.callLetters = "WDPG";
		station.frequency = "89.9";
		station.modulation = "FM";
		station.format = "Classical";
		station.networkOrOrganization = "Dayton Public Radio, Inc.";
		station.website = "https://discoverclassical.org/";
		stations.add(station);
		
		station = new RadioStation();
		station.callLetters = "WDAO";
		station.frequency = "1210";
		station.modulation = "AM";
		station.format = "Urban";
		station.networkOrOrganization = "";
		station.website = "https://wdaoradio.com/";
		stations.add(station);
		
		station = new RadioStation();
		station.callLetters = "WDAO";
		station.frequency = "102.3";
		station.modulation = "FM";
		station.format = "Urban";
		station.networkOrOrganization = "";
		station.website = "https://wdaoradio.com/";
		stations.add(station);
		
		station = new RadioStation();
		station.callLetters = "WLW";
		station.frequency = "700";
		station.modulation = "AM";
		station.format = "Sports";
		station.networkOrOrganization = "";
		station.website = "https://700wlw.iheart.com/";
		stations.add(station);
		
		station = new RadioStation();
		station.callLetters = "WKET";
		station.frequency = "98.3";
		station.modulation = "FM";
		station.format = "Rock";
		station.networkOrOrganization = "Kettering Fairmont";
		station.website = "";
		stations.add(station);
		
		station = new RadioStation();
		station.callLetters = "WUDR";
		station.frequency = "99.5";
		station.modulation = "FM";
		station.format = "College Radio";
		station.networkOrOrganization = "University of Dayton";
		station.website = "https://udayton.edu/studev/leadership/involvement/student-life/org-radio.php";
		stations.add(station);
		
		station = new RadioStation();
		station.callLetters = "WUDR";
		station.frequency = "98.1";
		station.modulation = "FM";
		station.format = "College Radio";
		station.networkOrOrganization = "University of Dayton";
		station.website = "https://udayton.edu/studev/leadership/involvement/student-life/org-radio.php";
		stations.add(station);
		
		station = new RadioStation();
		station.callLetters = "WWSU";
		station.frequency = "106.9";
		station.modulation = "FM";
		station.format = "College Radio";
		station.networkOrOrganization = "Wright State University";
		station.website = "https://www.wright.edu/wwsu";
		stations.add(station);
		
		station = new RadioStation();
		station.callLetters = "WVXU";
		station.frequency = "91.7";
		station.modulation = "FM";
		station.format = "News";
		station.networkOrOrganization = "NPR";
		station.website = "https://www.wvxu.org/";
		stations.add(station);
		
		station = new RadioStation();
		station.callLetters = "WDPS";
		station.frequency = "89.5";
		station.modulation = "FM";
		station.format = "Jazz";
		station.networkOrOrganization = "Dayton Public Schools";
		station.website = "https://www.dps.k12.oh.us/media-center/wdps-fm/";
		stations.add(station);
		
		station = new RadioStation();
		station.callLetters = "WZDA";
		station.frequency = "103.9";
		station.modulation = "FM";
		station.format = "Alternative rock";
		station.networkOrOrganization = "";
		station.website = "https://altdayton.iheart.com/";
		stations.add(station);
		
		station = new RadioStation();
		station.callLetters = "WTUE";
		station.frequency = "104.7";
		station.modulation = "FM";
		station.format = "Rock";
		station.networkOrOrganization = "";
		station.website = "https://wtue.iheart.com/";
		stations.add(station);
		
		station = new RadioStation();
		station.callLetters = "WMMX";
		station.frequency = "107.7";
		station.modulation = "FM";
		station.format = "Adult Contemporary";
		station.networkOrOrganization = "";
		station.website = "https://mix1077.iheart.com/";
		stations.add(station);
	}
	
	/**
	 * Defines the data type for a terrestrial radio station.
	 */
	private static class RadioStation {
		public String callLetters = "";
		public String frequency = "";
		public String modulation = "";
		public String format = "";
		public String networkOrOrganization = "";
		public String website = "";
	}
}
