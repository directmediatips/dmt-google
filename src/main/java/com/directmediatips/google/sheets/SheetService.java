package com.directmediatips.google.sheets;

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
import java.io.InputStream;
import java.io.InputStreamReader;
import java.util.Arrays;
import java.util.List;

import com.google.api.client.auth.oauth2.Credential;
import com.google.api.client.extensions.java6.auth.oauth2.AuthorizationCodeInstalledApp;
import com.google.api.client.extensions.jetty.auth.oauth2.LocalServerReceiver;
import com.google.api.client.googleapis.auth.oauth2.GoogleAuthorizationCodeFlow;
import com.google.api.client.googleapis.auth.oauth2.GoogleClientSecrets;
import com.google.api.client.googleapis.javanet.GoogleNetHttpTransport;
import com.google.api.client.http.HttpTransport;
import com.google.api.client.json.JsonFactory;
import com.google.api.client.json.jackson2.JacksonFactory;
import com.google.api.client.util.store.FileDataStoreFactory;
import com.google.api.services.sheets.v4.Sheets;
import com.google.api.services.sheets.v4.SheetsScopes;

/**
 * Service that will allow us to read from and write to a Google Sheets document.
 */
public class SheetService {
	
	/** The application name ("dmt-google"). */
	private static final String APPLICATION_NAME = "dmt-google";
	
	/** The path to the file "client_secret.json". */
	private static final String SECRET = "google/client_secret.json";
	
	/** The {@link JsonFactory} instance. */
	private static final JsonFactory JSON_FACTORY = JacksonFactory.getDefaultInstance();
	
	/** The scope is limited to Google sheets. */
	private static final List<String> SCOPES = Arrays.asList(SheetsScopes.SPREADSHEETS);
	
	/** The {@link HttpTransport} instance. */
	private static HttpTransport HTTP_TRANSPORT;
	
	/** The data store factory that will create a file named <code>StoredCredential</code> in a directory named "google". */
	private static FileDataStoreFactory DATA_STORE_FACTORY;
    static {
        try {
            HTTP_TRANSPORT = GoogleNetHttpTransport.newTrustedTransport();
            DATA_STORE_FACTORY = new FileDataStoreFactory(new java.io.File("google"));
        } catch (Throwable t) {
            t.printStackTrace();
            System.exit(1);
        }
    }
    
    /**
     * Creates a Credential object to get access to Google sheets.
     *
     * @return a Credential instance
     * @throws IOException Signals that an I/O exception has occurred.
     */
    public static Credential authorize() throws IOException {
        // Load client secrets.
        InputStream in =
            new FileInputStream(SECRET);
        GoogleClientSecrets clientSecrets =
            GoogleClientSecrets.load(JSON_FACTORY, new InputStreamReader(in));

        // Build flow and trigger user authorization request.
        GoogleAuthorizationCodeFlow flow =
                new GoogleAuthorizationCodeFlow.Builder(
                        HTTP_TRANSPORT, JSON_FACTORY, clientSecrets, SCOPES)
                .setDataStoreFactory(DATA_STORE_FACTORY)
                .setAccessType("offline")
                .build();
        Credential credential = new AuthorizationCodeInstalledApp(
                flow, new LocalServerReceiver()).authorize("twitter_app");
        return credential;
    }
    
    /**
     * Creates a Sheets service that will allow us to read and write from a Google Sheets document.
     *
     * @return a Sheets instance
     * @throws IOException Signals that an I/O exception has occurred.
     */
    public static Sheets getSheets() throws IOException {
        Credential credential = authorize();
        return new Sheets.Builder(HTTP_TRANSPORT, JSON_FACTORY, credential)
                .setApplicationName(APPLICATION_NAME)
                .build();
    }
}
