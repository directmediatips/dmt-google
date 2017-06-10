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
import com.google.api.services.sheets.v4.model.AddSheetRequest;
import com.google.api.services.sheets.v4.model.AddSheetResponse;
import com.google.api.services.sheets.v4.model.BatchUpdateSpreadsheetRequest;
import com.google.api.services.sheets.v4.model.BatchUpdateSpreadsheetResponse;
import com.google.api.services.sheets.v4.model.CellData;
import com.google.api.services.sheets.v4.model.ExtendedValue;
import com.google.api.services.sheets.v4.model.GridCoordinate;
import com.google.api.services.sheets.v4.model.Request;
import com.google.api.services.sheets.v4.model.Response;
import com.google.api.services.sheets.v4.model.RowData;
import com.google.api.services.sheets.v4.model.Sheet;
import com.google.api.services.sheets.v4.model.SheetProperties;
import com.google.api.services.sheets.v4.model.Spreadsheet;
import com.google.api.services.sheets.v4.model.UpdateCellsRequest;

/**
 * Will either get or create a sheet for a specific month and
 * a specific Twitter account in a Google Sheets document.
 */
public class TwitterMetrics {

	/** The Google Sheets service. */
	protected Sheets service;
	/** The ID of the spreadsheet with the Twitter information. */
	protected String spreadsheetId;
	/** The screen name of the Twitter account we're looking at. */
	protected String account;
	/** A date in the month for which we want a sheet of. */
	protected Date date;
	
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
	
	/**
	 * Creates a TwitterMetrics instance.
	 * @param service	the Sheets service
	 * @param spreadsheetId	an ID of a Google sheets document
	 * @param account	the screen name of a Twitter account
	 * @param date	a date (only the month is important)
	 */
	public TwitterMetrics(Sheets service, String spreadsheetId,
			String account, Date date) {
		this.service = service;
		this.spreadsheetId = spreadsheetId;
		this.account = account;
		this.date = date;
	}
	
	/**
	 * Formats the title of the sheet, using the Twitter account name,
	 * a year and a month.
	 *
	 * @return the sheet title
	 */
	public String getSheetTitle() {
		return String.format("%s-%s/%s",
			account,
			Integer.parseInt(new SimpleDateFormat("yyyy").format(date)),
	        new SimpleDateFormat("MMMM").format(date));
	}
	
	/**
	 * Gets the sheet id of a specific sheet in a Google Sheets document.
	 * If the sheet doesn't exist yet, a new sheet is created.
	 *
	 * @param sheetname the name of a sheet (see getSheetTitle())
	 * @return a sheet ID
	 * @throws IOException Signals that an I/O exception has occurred.
	 */
	public int getSheetId(String sheetname) throws IOException {
		// Gets the spreadsheet document
  	  	Spreadsheet sheet = service.spreadsheets().get(spreadsheetId).execute();
  	  	// Gets the sheets in the spreadsheet
  	  	List<Sheet> list = sheet.getSheets();
  	  	Integer id = null;
  	  	String title;
  	  	// Loops over the sheets looking for the sheet name
  	  	for (Sheet s : list) {
  	  		title = s.getProperties().getTitle();
  	  		if (sheetname.equals(title)) {
  	  			id = s.getProperties().getSheetId();
  	  			break;
  	  		}
  	  	}
  	  	// Create a new sheet if the sheet isn't found
  	  	if (id == null) {
  	  		id = createSheet(sheetname);
  	  		createHeader(id);
  	  	}
  	  	return id;
	}
	
	/**
	 * Creates a new sheet.
	 *
	 * @param title the title
	 * @return a sheet id
	 * @throws IOException Signals that an I/O exception has occurred.
	 */
	public int createSheet(String title) throws IOException {
	    List<Request> requests = new ArrayList<Request>();
	    requests.add(new Request()
	        .setAddSheet(new AddSheetRequest()
	        .setProperties(new SheetProperties().setTitle(title))
	    ));
	    BatchUpdateSpreadsheetRequest update =
	        new BatchUpdateSpreadsheetRequest().setRequests(requests);
	    BatchUpdateSpreadsheetResponse response =
	        service.spreadsheets().batchUpdate(spreadsheetId, update).execute();
	    List<Response> replies = response.getReplies();
	    AddSheetResponse sheetResponse = replies.get(0).getAddSheet();
	    return sheetResponse.getProperties().getSheetId();
	}
	
	/**
	 * Adds a header row to a specific sheet.
	 *
	 * @param sheetId the ID of the sheet
	 * @throws IOException Signals that an I/O exception has occurred.
	 */
	public void createHeader(int sheetId) throws IOException {
  	  	List<CellData> values = new ArrayList<CellData>();
  	  	values.add(new CellData().setUserEnteredValue(new ExtendedValue().setStringValue("Date")));
  	  	values.add(new CellData().setUserEnteredValue(new ExtendedValue().setStringValue("Tweets")));
  	  	values.add(new CellData().setUserEnteredValue(new ExtendedValue().setStringValue("Following")));
  	  	values.add(new CellData().setUserEnteredValue(new ExtendedValue().setStringValue("Followers")));
  	  	values.add(new CellData().setUserEnteredValue(new ExtendedValue().setStringValue("Likes")));
  	  	
  	  	List<Request> requests = new ArrayList<Request>();
  	  	UpdateCellsRequest updateCellRequest = new UpdateCellsRequest()
  	  			.setStart(new GridCoordinate()
  	  					.setSheetId(sheetId)
  	  					.setRowIndex(0)
  	  					.setColumnIndex(0))
  	  			.setRows(Arrays.asList(new RowData().setValues(values)))
  	  			.setFields("*");
  	  	requests.add(new Request().setUpdateCells(updateCellRequest));
  	  	
  	  	BatchUpdateSpreadsheetRequest batchUpdateRequest =
  	  			new BatchUpdateSpreadsheetRequest().setRequests(requests);
  	  	service.spreadsheets().batchUpdate(spreadsheetId, batchUpdateRequest).execute();
	}
}
