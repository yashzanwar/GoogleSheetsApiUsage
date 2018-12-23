import com.google.api.services.sheets.v4.model.ValueRange;
import org.testng.annotations.Test;

import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

/**
 * created by yash.zanwar on 21/12/18
 */
public class TestGoogleSheetApis {

    private static String spreasdsfID;


    @Test
    public void verifyProfileInfo() throws IOException {
        GoogleSheetApi sheetAPI = new GoogleSheetApi();
        spreasdsfID = sheetAPI.create("Mutual Funds");

        /** Updating the values in sheet **/
        Object[][] yash = {{"Month", "Mutual Funds", "Stocks", "Total"}, {"Jan", "200", "300", "322"}, {"Feb", "200", "332", "3443"}};
        List<List<Object>> writeData;
        writeData = convert2DArrayToListOfList(yash);
        sheetAPI.updateValues(spreasdsfID, "Sheet1!A1:D3", writeData);

        /**Reading from sheet **/
        ValueRange range = sheetAPI.getValues(spreasdsfID, "Sheet1!A1:D3");
        System.out.println(range.getValues().get(1).get(0));

        /**Formatting of cells **/
        sheetAPI.formattingCells(spreasdsfID);

        /** Making a line graph **/
        MakeCharts makeCharts = new MakeCharts();
        makeCharts.makeChart(spreasdsfID, "title");


    }

    public List<List<Object>> convert2DArrayToListOfList(Object[][] array) {
        List<List<Object>> writeData = new ArrayList<List<Object>>();
        for (Object[] someData : array) {
            List<Object> dataRow = new ArrayList<Object>();
            for (int i = 0; i < array[0].length; i++) {
                dataRow.add(someData[i]);
            }
            writeData.add(dataRow);
        }
        return writeData;
    }

}
