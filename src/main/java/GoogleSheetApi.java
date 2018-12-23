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
import com.google.api.services.sheets.v4.model.*;

import java.io.IOException;
import java.io.InputStream;
import java.io.InputStreamReader;
import java.util.ArrayList;
import java.util.Collections;
import java.util.List;


/**
 * created by yash.zanwar on 21/12/18
 */
public class GoogleSheetApi {
    private static final String APPLICATION_NAME = "Google Sheets API Java Quickstart";
    private static final JsonFactory JSON_FACTORY = JacksonFactory.getDefaultInstance();

    /**
     * Global instance of the scopes required by this quickstart.
     * If modifying these scopes, delete your previously saved tokens/ folder.
     */
    private static final List<String> SCOPES = Collections.singletonList(SheetsScopes.SPREADSHEETS);
    private static final String CREDENTIALS_FILE_PATH = "/credentials.json";
    /**
     * Directory to store user credentials for this application.
     */
    private static final java.io.File DATA_STORE_DIR = new java.io.File(
            System.getProperty("user.home"), ".credentials/sheets.googleapis.com-java-quickstart");

    private static FileDataStoreFactory DATA_STORE_FACTORY;

    private static HttpTransport HTTP_TRANSPORT;

    static {
        try {
            HTTP_TRANSPORT = GoogleNetHttpTransport.newTrustedTransport();
            DATA_STORE_FACTORY = new FileDataStoreFactory(DATA_STORE_DIR);
        } catch (Throwable t) {
            t.printStackTrace();
            System.exit(1);
        }
    }


    /**
     * Creates an authorized Credential object.
     *
     * @return An authorized Credential object.
     * @throws IOException If the credentials.json file cannot be found.
     */
    public static Credential authorize() throws IOException {
        // Load client secrets.
        InputStream in = GoogleSheetApi.class.getResourceAsStream(CREDENTIALS_FILE_PATH);
        GoogleClientSecrets clientSecrets = GoogleClientSecrets.load(JSON_FACTORY, new InputStreamReader(in));

        // Build flow and trigger user authorization request.
        GoogleAuthorizationCodeFlow flow = new GoogleAuthorizationCodeFlow.Builder(HTTP_TRANSPORT, JSON_FACTORY, clientSecrets, SCOPES)
                .setDataStoreFactory(DATA_STORE_FACTORY)
                .setAccessType("offline")
                .build();
        Credential credential = new AuthorizationCodeInstalledApp(flow, new LocalServerReceiver()).authorize("user");
        System.out.println("Credentials saved to " + DATA_STORE_DIR.getAbsolutePath());
        return credential;
    }

    /**
     * Build and return an authorized Sheets API client service.
     *
     * @return an authorized Sheets API client service
     * @throws IOException
     */
    public static Sheets getSheetsService() throws IOException {
        Credential credential = authorize();
        return new Sheets.Builder(HTTP_TRANSPORT, JSON_FACTORY, credential)
                .setApplicationName(APPLICATION_NAME)
                .build();
    }

    public List<List<Object>> getSpreadSheetRecords(String spreadsheetId, String range) throws IOException {
        Sheets service = getSheetsService();
        ValueRange response = service.spreadsheets().values()
                .get(spreadsheetId, range)
                .execute();
        List<List<Object>> values = response.getValues();
        if (values != null && values.size() != 0) {
            return values;
        } else {
            System.out.println("No data found.");
            return null;
        }
    }

    public String create(String title) throws IOException {
        Sheets service = getSheetsService();
        Spreadsheet spreadsheet = new Spreadsheet()
                .setProperties(new SpreadsheetProperties()
                        .setTitle(title));
        spreadsheet = service
                .spreadsheets()
                .create(spreadsheet)
                .setFields("spreadsheetId")
                .execute();
        System.out.println("Spreadsheet ID: " + spreadsheet.getSpreadsheetId());
        return spreadsheet.getSpreadsheetId();
    }


    public UpdateValuesResponse updateValues(String spreadsheetId, String range, List<List<Object>> _values) throws IOException {
        Sheets service = getSheetsService();
        ValueRange vr = new ValueRange()
                .setValues(_values)
                .setMajorDimension("ROWS");
        UpdateValuesResponse result = service
                .spreadsheets()
                .values()
                .update(spreadsheetId, range, vr)
                .setValueInputOption("USER_ENTERED").execute();
        System.out.printf("%d cells updated.", result.getUpdatedCells());


        return result;
    }

    public ValueRange getValues(String spreadsheetId, String range) throws IOException {
        Sheets service = getSheetsService();
        ValueRange result = service
                .spreadsheets()
                .values()
                .get(spreadsheetId, range)
                .execute();
        int numRows = result.getValues() != null ? result.getValues().size() : 0;
        System.out.printf("%d rows retrieved.", numRows);
        return result;
    }

    public BatchUpdateSpreadsheetResponse formattingCells(String spreadsheetId) throws IOException {
        Sheets service = getSheetsService();
        List<Request> requests = new ArrayList<Request>();

        //Formatting of the borders
        Border borderTop = new Border()
                .setStyle("DASHED")
                .setWidth(10)
                .setColor(new Color()
                        .setRed(1f)
                        .setGreen(0.f)
                        .setBlue(0.f));
        Border borderBottom = new Border()
                .setStyle("DASHED")
                .setWidth(10).setColor(new Color()
                        .setRed(1f)
                        .setGreen(0f)
                        .setBlue(0f));
        GridRange tableRange = setRangeOfCells(0, 1, 10, 1, 10);

        requests.add(new Request()
                .setUpdateBorders(new UpdateBordersRequest()
                        .setRange(tableRange)
                        .setTop(borderTop)
                        .setBottom(borderBottom)));

        //Formatting of the cells
        tableRange = setRangeOfCells(0, 0, 1, 0, 4);
        CellData cellData = new CellData()
                .setUserEnteredFormat(new CellFormat()
                        .setBackgroundColor(new Color().setRed(0f).setGreen(0f).setBlue(0f))
                        .setHorizontalAlignment("CENTER")
                        .setTextFormat(new TextFormat()
                                .setForegroundColor(new Color().setBlue(1f).setRed(1f).setGreen(1f))
                                .setFontSize(12)
                                .setBold(true)));

        SheetProperties sheetProperties = new SheetProperties().setSheetId(0).setGridProperties(new GridProperties().setFrozenRowCount(1));
        requests.add(new Request()
                .setRepeatCell(new RepeatCellRequest()
                        .setRange(tableRange)
                        .setCell(cellData)
                        .setFields("userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)")));

        //Freezing of header row
        requests.add(new Request()
                .setUpdateSheetProperties(new UpdateSheetPropertiesRequest()
                        .setProperties(sheetProperties)
                        .setFields("gridProperties.frozenRowCount")));

        //Merging of Cells
        tableRange = setRangeOfCells(0, 4, 6, 0, 3);
        requests.add(new Request()
                .setMergeCells(new MergeCellsRequest()
                        .setRange(tableRange)
                        .setMergeType("MERGE_ALL")));
        tableRange = setRangeOfCells(0, 4, 8, 5, 9);
        requests.add(new Request()
                .setMergeCells(new MergeCellsRequest()
                        .setRange(tableRange)
                        .setMergeType("MERGE_COLUMNS")));

        //Setting Date format
        tableRange = setRangeOfCells(0, 0, 100, 0, 100);
        cellData = new CellData()
                .setUserEnteredFormat(new CellFormat()
                        .setNumberFormat(new NumberFormat()
                                .setType("DATE")
                                .setPattern("hh:mm:ss am/pm, ddd mmm dd yyyy")));
        requests.add(new Request()
                .setRepeatCell(new RepeatCellRequest()
                        .setRange(tableRange)
                        .setCell(cellData)
                        .setFields("userEnteredFormat.numberFormat")));

        //Setting Decimal Format
        tableRange = setRangeOfCells(0, 1, 3, 1, 3);
        cellData = new CellData()
                .setUserEnteredFormat(new CellFormat()
                        .setNumberFormat(new NumberFormat()
                                .setType("NUMBER")
                                .setPattern("#,##0.0000")));
        requests.add(new Request()
                .setRepeatCell(new RepeatCellRequest()
                        .setRange(tableRange).setCell(cellData)
                        .setFields("userEnteredFormat.numberFormat")));

        BatchUpdateSpreadsheetRequest body = new BatchUpdateSpreadsheetRequest().setRequests(requests);
        BatchUpdateSpreadsheetResponse response = service.spreadsheets().batchUpdate(spreadsheetId, body).execute();
        System.out.printf("%d cells updated.", response.getReplies().size());
        return response;
    }

    public GridRange setRangeOfCells(int sheetId, int startRow, int endRow, int startColumn, int endColumn) {
        GridRange tableRange = new GridRange()
                .setSheetId(sheetId)
                .setStartRowIndex(startRow)
                .setEndRowIndex(endRow)
                .setStartColumnIndex(startColumn)
                .setEndColumnIndex(endColumn);
        return tableRange;
    }
}
