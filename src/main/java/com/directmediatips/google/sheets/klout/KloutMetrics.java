package com.directmediatips.google.sheets.klout;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Calendar;
import java.util.Date;
import java.util.List;
import java.util.Properties;

import com.directmediatips.google.sheets.SheetService;
import com.directmediatips.google.sheets.twitter.AbstractTwitterSheet;
import com.directmediatips.google.sheets.twitter.TwitterMetrics;
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
 * allow storing the Klout score.
 */
public class KloutMetrics extends AbstractTwitterSheet {

	/**
	 * Creates a KloutMetrics instance.
	 * @param service	the Sheets service
	 * @param spreadsheetId	an ID of a Google sheets document
	 * @param account	the screen name of a Twitter account
	 * @param date	a date (only the month is important)
	 */
	public KloutMetrics(Sheets service, String spreadsheetId,
			String account, Date date) {
		super(service, spreadsheetId, account, date);
	}
	
	/**
	 * Updates the metrics of a Twitter account in a Google spreadsheet.
	 *
	 * @param account the screenname of a Twitter account
	 * @param score	the current Klout score
	 * @throws IOException Signals that an I/O exception has occurred.
	 */
	public static void UpdateMetrics(String account, double score) throws IOException {
		// Initialize the TwitterMetrics instance
  	  	Sheets service = SheetService.getSheets();
		Properties props = new Properties();
		props.load(new FileInputStream("google/sheet.properties"));
		String spreadsheetId = props.getProperty("twitterMetrics");
  	  	Date date = new Date(); 
		TwitterMetrics sheetname = new TwitterMetrics(service, spreadsheetId, account, date);
		// Initialize the cell data
  	  	List<CellData> values = new ArrayList<CellData>();
  	  	values.add(new CellData().setUserEnteredValue(new ExtendedValue().setNumberValue(score)));
  	  	// Create the update request
  	    Calendar cal = Calendar.getInstance();
  	    cal.setTime(date);
  	  	List<Request> requests = new ArrayList<Request>();
  	  	UpdateCellsRequest updateCellRequest = new UpdateCellsRequest()
  	  			.setStart(new GridCoordinate()
  	  					.setSheetId(sheetname.getSheetId(sheetname.getSheetTitle()))
  	  					.setRowIndex(cal.get(Calendar.DAY_OF_MONTH))
  	  					.setColumnIndex(6))
  	  			.setRows(Arrays.asList(new RowData().setValues(values)))
  	  			.setFields("*");
  	  	requests.add(new Request().setUpdateCells(updateCellRequest));
  	  	// Execute the request
  	  	BatchUpdateSpreadsheetRequest batchUpdateRequest =
  	  			new BatchUpdateSpreadsheetRequest().setRequests(requests);
  	  	service.spreadsheets().batchUpdate(spreadsheetId, batchUpdateRequest).execute();
	}
}
