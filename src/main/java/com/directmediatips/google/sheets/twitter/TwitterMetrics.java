package com.directmediatips.google.sheets.twitter;

import java.io.FileInputStream;

/*
 * Copyright 2017, Bruno Lowagie, Wil-Low BVBA
 * 
 * Licensed under the Apache License, Version 2.0 (the "License"); you may
 * not use this file except in compliance with the License. You may obtain
 * a copy of the License at http://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the  * specific language governing permissions and
 * limitations under the License.
 */

import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Calendar;
import java.util.Date;
import java.util.List;
import java.util.Properties;

import com.directmediatips.google.sheets.SheetService;
import com.google.api.services.sheets.v4.Sheets;
import com.google.api.services.sheets.v4.model.BatchUpdateSpreadsheetRequest;
import com.google.api.services.sheets.v4.model.CellData;
import com.google.api.services.sheets.v4.model.ExtendedValue;
import com.google.api.services.sheets.v4.model.GridCoordinate;
import com.google.api.services.sheets.v4.model.Request;
import com.google.api.services.sheets.v4.model.RowData;
import com.google.api.services.sheets.v4.model.UpdateCellsRequest;

/**
 * Will either get or create a sheet for a specific month and
 * a specific Twitter account in a Google Sheets document to
 * allow storing Twitter metrics such as number of tweets,
 * number of friends, number of followers, number of likes.
 */
public class TwitterMetrics extends AbstractTwitterSheet {
	
	/**
	 * Creates a TwitterMetrics instance.
	 * @param service	the Sheets service
	 * @param spreadsheetId	an ID of a Google sheets document
	 * @param account	the screen name of a Twitter account
	 * @param date	a date (only the month is important)
	 */
	public TwitterMetrics(Sheets service, String spreadsheetId,
			String account, Date date) {
		super(service, spreadsheetId, account, date);
	}
	
	/**
	 * Updates the metrics of a Twitter account in a Google spreadsheet.
	 *
	 * @param account the screenname of a Twitter account
	 * @param tweets the number of tweets
	 * @param following the number of friends
	 * @param followers the number of followers
	 * @param likes 	the number of likes
	 * @throws IOException Signals that an I/O exception has occurred.
	 */
	public static void UpdateMetrics(String account,
			int tweets, int following, int followers, int likes) throws IOException {
		// Initialize the TwitterMetrics instance
  	  	Sheets service = SheetService.getSheets();
		Properties props = new Properties();
		props.load(new FileInputStream("google/sheet.properties"));
		String spreadsheetId = props.getProperty("twitterMetrics");
  	  	Date date = new Date(); 
		TwitterMetrics sheetname = new TwitterMetrics(service, spreadsheetId, account, date);
		// Initialize the cell data
  	  	List<CellData> values = new ArrayList<CellData>();
  	  	values.add(new CellData().setUserEnteredValue(new ExtendedValue().setStringValue(new SimpleDateFormat("yyyy-MM-dd").format(date))));
  	  	values.add(new CellData().setUserEnteredValue(new ExtendedValue().setNumberValue((double)tweets)));
  	  	values.add(new CellData().setUserEnteredValue(new ExtendedValue().setNumberValue((double)following)));
  	  	values.add(new CellData().setUserEnteredValue(new ExtendedValue().setNumberValue((double)followers)));
  	  	values.add(new CellData().setUserEnteredValue(new ExtendedValue().setNumberValue((double)likes)));
  	  	// Create the update request
  	    Calendar cal = Calendar.getInstance();
  	    cal.setTime(date);
  	  	List<Request> requests = new ArrayList<Request>();
  	  	UpdateCellsRequest updateCellRequest = new UpdateCellsRequest()
  	  			.setStart(new GridCoordinate()
  	  					.setSheetId(sheetname.getSheetId(sheetname.getSheetTitle()))
  	  					.setRowIndex(cal.get(Calendar.DAY_OF_MONTH))
  	  					.setColumnIndex(0))
  	  			.setRows(Arrays.asList(new RowData().setValues(values)))
  	  			.setFields("*");
  	  	requests.add(new Request().setUpdateCells(updateCellRequest));
  	  	// Execute the request
  	  	BatchUpdateSpreadsheetRequest batchUpdateRequest =
  	  			new BatchUpdateSpreadsheetRequest().setRequests(requests);
  	  	service.spreadsheets().batchUpdate(spreadsheetId, batchUpdateRequest).execute();
	}
}
