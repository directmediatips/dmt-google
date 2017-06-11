package com.directmediatips.google.sheets.twitter;

import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Date;
import java.util.List;

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

public class AbstractTwitterSheet {
	
	/** The Google Sheets service. */
	protected Sheets service;
	/** The ID of the spreadsheet with the Twitter information. */
	protected String spreadsheetId;
	/** The screen name of the Twitter account we're looking at. */
	protected String account;
	/** A date in the month for which we want a sheet of. */
	protected Date date;
	
	/**
	 * Creates an AbstractSheetHandler instance.
	 * @param service	the Sheets service
	 * @param spreadsheetId	an ID of a Google sheets document
	 * @param account	the screen name of a Twitter account
	 * @param date	a date (only the month is important)
	 */
	public AbstractTwitterSheet(Sheets service, String spreadsheetId,
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
  	  		if (sheetname.equalsIgnoreCase(title)) {
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
  	  	values.add(new CellData().setUserEnteredValue(new ExtendedValue().setStringValue("")));
  	  	values.add(new CellData().setUserEnteredValue(new ExtendedValue().setStringValue("Klout score")));
  	  	
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
