package com.directmediatips.google.sheets.twitter;

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

import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;
import java.util.Properties;
import java.util.TreeMap;

import com.directmediatips.google.sheets.SheetService;
import com.google.api.services.sheets.v4.Sheets;
import com.google.api.services.sheets.v4.model.ValueRange;

/**
 * Uses a Google spreadsheet to list who is following us based on search criteria.
 */
public class TwitterRichData {

	/**
	 * Inner class to store account info.
	 */
	public class Account {
		
		/** The screen name. */
		public String screenname;
		
		/** The accounts of interest. */
		public List<Integer> accounts = new ArrayList<Integer>();
	}
	
	/** An ordered map with account that follow at least one of our accounts. */
	protected Map<Long, Account> accounts = new TreeMap<Long, Account>();

	/** The range where we can find the WHERE clause in the Google spreadsheet. */
	public static final String RANGE1 = "criteria!A1";
	/** The range where we can find our twitter accounts in the Google spreadsheet. */
	public static final String RANGE2 = "results!C1:Z";
	/** The range where to put the results in the Google spreadsheet. */
	public static final String RANGE3 = "results!A2:Z";
	/** The range where we can find the direct message in the Google spreadsheet. */
	public static final String RANGE4 = "mail!A1";
	
	/** The Google Sheets service. */
	protected Sheets service;
	/** The ID of the spreadsheet with the Twitter information. */
	protected String spreadsheetId;
	
	/**
	 * Creates a TwitterData instance.
	 *
	 * @throws IOException Signals that an I/O exception has occurred.
	 */
	public TwitterRichData() throws IOException {
		this.service = SheetService.getSheets();
		Properties props = new Properties();
		props.load(new FileInputStream("google/sheet.properties"));
		this.spreadsheetId = props.getProperty("twitterRichData");
	}
	
	/**
	 * Gets the part of an SQL statement that defines the criteria. 
	 *
	 * @return the part of the query that comes after WHERE
	 * @throws IOException Signals that an I/O exception has occurred.
	 */
	public String getWhereClause() throws IOException {
		ValueRange response = service
				.spreadsheets()
				.values()
				.get(spreadsheetId, RANGE1)
				.execute();
		List<List<Object>> values = response.getValues();
		return values.get(0).get(0).toString();
	}

	/**
	 * Gets a list of Twitter accounts from the spreadsheet.
	 *
	 * @return a List of Twitter accounts
	 * @throws IOException Signals that an I/O exception has occurred.
	 */
	public List<Object> getFromAccounts() throws IOException {
		ValueRange response = service
				.spreadsheets()
				.values()
				.get(spreadsheetId, RANGE2)
				.execute();
		List<List<Object>> values = response.getValues();
		return values.get(0);
	}
	
	/**
	 * Adds an account to our list of accounts.
	 *
	 * @param id the account id
	 * @param screenname the account's screen name
	 * @param count an index referring to one of our own accounts
	 */
	public void add(long id, String screenname, int count) {
		Account account = accounts.get(id);
		if (account == null) {
			account = new Account();
			account.screenname = screenname;
		}
		account.accounts.add(count);
		accounts.put(id, account);
	}
	
	/**
	 * Processes the list of accounts, and writes the results to the Google sheets document.
	 *
	 * @param count the total number of accounts on our end
	 * @throws IOException Signals that an I/O exception has occurred.
	 */
	public void process(int count) throws IOException {
		List<List<Object>> data = new ArrayList<List<Object>>();
		List<Object> row;
		for (Map.Entry<Long, Account> entry : accounts.entrySet()) {
			row = new ArrayList<Object>();
			row.add(entry.getKey().toString());
			Account account = entry.getValue();
			row.add(account.screenname);
			for (int i = 0; i < count; i++) {
				if (account.accounts.contains(i)) {
					row.add(1);
				}
				else {
					row.add(0);
				}
			}
			data.add(row);
		}
		update(data);
	}
	
	/**
	 * Updates the spreadsheet with the data as stored in the database.
	 *
	 * @param data a two-dimensional array with harvest data as stored in the database
	 * @throws IOException Signals that an I/O exception has occurred.
	 */
	public void update(List<List<Object>> data) throws IOException {
		ValueRange valueRange = new ValueRange();
        valueRange.setValues(data);
        service.spreadsheets()
        	.values()
        	.update(spreadsheetId, RANGE3, valueRange)
            .setValueInputOption("RAW")
            .execute();
	}
	
	/**
	 * Gets the direct message from the Google spreadsheet. 
	 *
	 * @return the message
	 * @throws IOException Signals that an I/O exception has occurred.
	 */
	public String getDirectMessage() throws IOException {
		ValueRange response = service
				.spreadsheets()
				.values()
				.get(spreadsheetId, RANGE4)
				.execute();
		List<List<Object>> values = response.getValues();
		return values.get(0).get(0).toString();
	}
	
	/**
	 * Reads followers info from the Google spreadsheet.
	 *
	 * @return a Map with account IDs as key and extra info as value.
	 * @throws IOException Signals that an I/O exception has occurred.
	 */
	public Map<Long, Account> getToAccounts() throws IOException {
		Map<Long, Account> map = new TreeMap<Long, Account>();
		ValueRange response = service.spreadsheets().values()
	            .get(spreadsheetId, RANGE3)
	            .execute();
		List<List<Object>> values = response.getValues();
		if (values != null && values.size() > 0) {
			long id;
			for (List<Object> row : values) {
				id = getId(row.get(0).toString());
				if (id > 0)
					map.put(id,	getAccount(row));
			}
		}
		return map;
	}
	
	/**
	 * Converts a String into a long, but returns a unique negative
	 * number if not successful. Most of the times the String will be
	 * a number, but when new accounts are added, the String will be
	 * empty, and we'll want the system to find out the Twitter id.
	 * @param id	the String with the id
	 * @return	a long value
	 */
	protected long getId(String id) {
		try {
			return Long.parseLong(id);
		}
		catch(NumberFormatException nfe) {
			return 0;
		}
	}
	
	/**
	 * Converts a row of harvest data into an Account object.
	 * @param row	a row of data obtained from a Google spreadsheet
	 * @return	an Account object
	 */
	protected Account getAccount(List<Object> row) {
		Account account = new Account();
		account.screenname = (String)row.get(1);
		for (int i = 2; i < row.size(); i++) {
			if ("1".equals(row.get(i).toString()))
				account.accounts.add(i - 2);
		}
		return account;
	}
}
