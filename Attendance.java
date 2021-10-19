// 1. Two json files.  Either one works
// 2. Quotas for write requests per minute per user (Code: 429)
// 3. Delete token files in folder if scopes change
// 4. Occasional read timeout
// 5. row (worked) or column (?)
// a. The cell values are "Object" List<List<Object>> values = Arrays.asList(Arrays.asList(newRow));
// b.
import com.google.api.client.auth.oauth2.Credential;
import com.google.api.client.extensions.java6.auth.oauth2.AuthorizationCodeInstalledApp;
import com.google.api.client.extensions.jetty.auth.oauth2.LocalServerReceiver;
import com.google.api.client.googleapis.auth.oauth2.GoogleAuthorizationCodeFlow;
import com.google.api.client.googleapis.auth.oauth2.GoogleClientSecrets;
import com.google.api.client.googleapis.javanet.GoogleNetHttpTransport;
import com.google.api.client.http.javanet.NetHttpTransport;
import com.google.api.client.json.JsonFactory;
import com.google.api.client.json.jackson2.JacksonFactory;
import com.google.api.client.util.store.FileDataStoreFactory;
import com.google.api.services.sheets.v4.Sheets;
import com.google.api.services.sheets.v4.SheetsScopes;
import com.google.api.services.sheets.v4.model.ValueRange;
import com.google.api.services.sheets.v4.model.UpdateValuesResponse;

import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStream;
import java.io.InputStreamReader;
import java.security.GeneralSecurityException;
import java.util.Collections;
import java.util.List;
import java.util.Arrays;
import java.lang.System;

public class Attendance {
    private static final String APPLICATION_NAME = "Attendance";
    private static final JsonFactory JSON_FACTORY = JacksonFactory.getDefaultInstance();
    private static final String TOKENS_DIRECTORY_PATH = "tokens";

    /** Global instance of the scopes required by this quickstart.
     * If modifying these scopes, delete your previously saved tokens/ folder */
    private static final List<String> SCOPES = Collections.singletonList(SheetsScopes.SPREADSHEETS);
    private static final String CREDENTIALS_FILE_PATH = "/cred-desktop.json";
    /** Creates an authorized Credential object.
     * @param HTTP_TRANSPORT The network HTTP Transport.
     * @return An authorized Credential object.
     * @throws IOException If the credentials.json file cannot be found. */
    private static Credential getCredentials(final NetHttpTransport HTTP_TRANSPORT) throws IOException {
        // Load client secrets.
        InputStream in = Attendance.class.getResourceAsStream(CREDENTIALS_FILE_PATH);
        if (in == null) {
            throw new FileNotFoundException("Resource not found: " + CREDENTIALS_FILE_PATH);
        }
        GoogleClientSecrets clientSecrets = GoogleClientSecrets.load(JSON_FACTORY, new InputStreamReader(in));
        // Build flow and trigger user authorization request.
        GoogleAuthorizationCodeFlow flow = new GoogleAuthorizationCodeFlow.Builder(
                HTTP_TRANSPORT, JSON_FACTORY, clientSecrets, SCOPES)
                .setDataStoreFactory(new FileDataStoreFactory(new java.io.File(TOKENS_DIRECTORY_PATH)))
                .setAccessType("offline")
                .build();
        LocalServerReceiver receiver = new LocalServerReceiver.Builder().setPort(8888).build();
        return new AuthorizationCodeInstalledApp(flow, receiver).authorize("user");
    }

    public static void main(String... args) throws IOException, GeneralSecurityException {
        // Build a new authorized API client service.
        final NetHttpTransport HTTP_TRANSPORT = GoogleNetHttpTransport.newTrustedTransport();
        final String urlId = "1Lij9_onCKsr7CWVI4-GKEz73XMwEXljDsUP6aqcJEV0";   // URL of the spreadsheet
        final String attsheet = "All Team 9/17";  // Attendance spreadsheet name 
        final String rossheet = "All Team";       // Roster spreadsheet name
        final String attrange = attsheet + "!A1:D100";           //
        final String rosrange = rossheet + "!A1:X100";           // The whole range possible
        String valueInputOption = "USER_ENTERED";        // Determines how input data should be interpreted. RAW
//        ValueRange titleBody = new ValueRange().setValues(Arrays.asList(Arrays.asList(attsheet))); // 
//        ValueRange requestBody = new ValueRange().setValues(Arrays.asList(Arrays.asList("X")));    // 

        Sheets service = new Sheets.Builder(HTTP_TRANSPORT, JSON_FACTORY, getCredentials(HTTP_TRANSPORT))
                .setApplicationName(APPLICATION_NAME).build();
        // Retrieve attendance names
        ValueRange attlists = service.spreadsheets().values().get(urlId, attrange).execute();
        ValueRange roslists = service.spreadsheets().values().get(urlId, rosrange).execute();
        List<List<Object>> attnames = attlists.getValues();
        List<List<Object>> rosnames = roslists.getValues();
        int attnum = attnames.size() - 1;                  // Row 1 is title
        int rosnum = rosnames.size() - 1;                  // Row 1 is title
        int chkdnum = 0;                                    // Number of people checked in roster
        int intrdnum = 0;
        int newcolascii = (int) ('A' + rosnames.get(0).size());  // Creating A1 notation
        String newcol = Character.toString((char) newcolascii);  // New column to add
        Object[][] newColumn = new Object[rosnum+1][1];             // Title + All the names
        Object[]newRow       = new Object[rosnum+1];             // Title + All the names
        newColumn[0][0] = attsheet;
        newRow[0] = attsheet;
        String attrosrange;
        for (int k=1; k < rosnum+1; k++) {
            newColumn[k][0] = "";
            newRow[k] = "";
        }

        if (attnames == null || attnames.isEmpty()) {       // Exit if cannot find Attendance
            System.out.println("No data found in Attendance spreadsheet.");
            System.exit(0);
        } 
        Boolean found = false;
        for (int i=1; i<=attnum; i++) {                // Check every name in the attendance list
            List attrow = attnames.get(i);
            String attfirst = attrow.get(2).toString().toLowerCase().trim();
            String attlast  = attrow.get(1).toString().toLowerCase().trim();
            int j = 1; found = false;
            while (j<=rosnum && !found) {
                List rosrow = rosnames.get(j);         // Compare with every name in the roster
                String rosfirst = rosrow.get(1).toString().toLowerCase().trim();
                String roslast = rosrow.get(0).toString().toLowerCase().trim();
                String rosnick;
                int roslen = rosrow.size();
                if (roslast.contains(attlast) || attlast.contains(roslast)) {      // Matched Last Name
                    if (roslen <= 2) {
                        rosnick = "XXXX";                    // Not matchable character
                    } else if (rosrow.get(2) == null) {
                        rosnick = "XXX";
                    } else if (rosrow.get(2) == "" || rosrow.get(2) == " ") {
                        rosnick = "XX";
                    } else {
                        rosnick = rosrow.get(2).toString().toLowerCase().trim();
                    }
                    if (rosfirst.contains(attfirst) || attfirst.contains(rosfirst) || rosnick.contains(attfirst) || attfirst.contains(rosnick)) {
                        System.out.printf("%d, %d: %s %s %s (%s) %s\n", i, j, attfirst, attlast, rosfirst, rosnick, roslast);
//UpdateValuesResponse result = service.spreadsheets().values().update(urlId, attrosrange, requestBody).setValueInputOption(valueInputOption).execute();
// Sheets.Spreadsheets.Values.Update request =  service.spreadsheets().values().update(urlId, attrosrange, requestBody);
// request.setValueInputOption(valueInputOption);
// UpdateValuesResponse response = request.execute();
                        newColumn[j][0] = "X";
                        newRow[j] = "X";
                        chkdnum++;
                        found = true;
                    }     
                }   
                j++;
            }  // for j
            if (!found) {
                System.out.printf("%d -> Intruder %s %s\n", i, attfirst, attlast);
                intrdnum++;
            }
        }   // for i 
        attrosrange = rossheet + "!" + newcol + "1:" + newcol + String.valueOf(rosnum+1);   // The Cell in Alphabet and Number A1:N100
//  attrosrange = rossheet + "!" + newcol + "1";   // The Cell in Alphabet and Number A1:N100
        System.out.printf("%d in roster, %d came, %d checked, %d intruded: added to %s\n", rosnum, attnum, chkdnum, intrdnum, attrosrange);
        for (int k=0; k<=rosnum; k++) {
            System.out.print(newColumn[k][0]);
            System.out.print(" ");
        }
//        List<List<Object>> values = Arrays.asList(Arrays.asList(newColumn));
        List<List<Object>> valuelists = Arrays.asList(Arrays.asList(newRow));
        ValueRange updateBody = new ValueRange().setValues(valuelists).setMajorDimension("COLUMNS").setRange(attrosrange);
        UpdateValuesResponse result = service.spreadsheets().values()
                .update(urlId, attrosrange, updateBody)
                .setValueInputOption(valueInputOption)
                .execute();
    }  //main()
}
//Sheets.Spreadsheets.Values.Update request =  service.spreadsheets().values().update(urlId, attrosrange, updateBody);
//request.setValueInputOption(valueInputOption);
//UpdateValuesResponse result = request.execute();
//try {
//    // thread to sleep for 1000 milliseconds
//    Thread.sleep(60000);
// } catch (Exception e) {}
