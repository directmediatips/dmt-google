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
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.Properties;

import com.directmediatips.google.sheets.SheetService;
import com.google.api.services.sheets.v4.Sheets;
import com.google.api.services.sheets.v4.model.ValueRange;

/**
 * Uses a Google spreadsheet to manage harvest account info.
 */
public class TwitterHarvest {

	/**
	 * Inner class to store harvest info.
	 */
	public class Account {
		
		/** The screen name. */
		public String screenname;
		
		/** The follow friends. */
		public int followFriends;
		
		/** The follow followers. */
		public int followFollowers;
		
		/** The retweet. */
		public int retweet;
		
		/**
		 * Checks if is unchanged.
		 *
		 * @param f1 the follow friends value
		 * @param f2 the follow followers value
		 * @param r the retweet value
		 * @return true, if none of the values changed
		 */
		public boolean isUnchanged(int f1, int f2, int r) {
			return f1 == followFriends && f2 == followFollowers && r == retweet;
		}
	}
	
	/** The range where we can find harvest info in the Google spreadsheet. */
	public static final String RANGE = "%s!A1:E";
	
	/** The Google Sheets service. */
	protected Sheets service;
	/** The ID of the spreadsheet with the Twitter information. */
	protected String spreadsheetId;
	/** The screen name of the Twitter account we're looking at. */
	protected String account;
	
	/** Counter keeping track of new accounts. */
	protected long count = 0;
	
	/**
	 * Creates a TwitterHarvest instance.
	 *
	 * @param account the screen name of a Twitter account
	 * @throws IOException Signals that an I/O exception has occurred.
	 */
	public TwitterHarvest(String account) throws IOException {
		this.service = SheetService.getSheets();
		Properties props = new Properties();
		props.load(new FileInputStream("google/sheet.properties"));
		this.spreadsheetId = props.getProperty("twitterHarvest");
		this.account = account;
	}
	
	/**
	 * Reads Harvest info from the Google spreadsheet.
	 *
	 * @return a Map with account IDs as key and harvest info as value.
	 * @throws IOException Signals that an I/O exception has occurred.
	 */
	public Map<Long, Account> getHarvestData() throws IOException {
		Map<Long, Account> map = new HashMap<Long, Account>();
		ValueRange response = service.spreadsheets().values()
	            .get(spreadsheetId, String.format(RANGE, account))
	            .execute();
		List<List<Object>> values = response.getValues();
		if (values != null && values.size() > 0) {
			for (List<Object> row : values) {
				map.put(getId(row.get(0).toString()),
						getAccount(row));
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
			return --count;
		}
	}
	
	/**
	 * Converts a row of harvest data into an Account object.
	 * @param row	a row of data obtained from a Google spreadsheet
	 * @return	an Account object
	 */
	protected Account getAccount(List<Object> row) {
		Account harvest = new Account();
		harvest.screenname = (String)row.get(1);
		harvest.followFriends = Integer.parseInt(row.get(2).toString());
		harvest.followFollowers = Integer.parseInt(row.get(3).toString());
		harvest.retweet = Integer.parseInt(row.get(4).toString());
		return harvest;
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
        	.update(spreadsheetId, String.format(RANGE, account), valueRange)
            .setValueInputOption("RAW")
            .execute();
	}
}
