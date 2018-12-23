import com.google.api.services.sheets.v4.Sheets;
import com.google.api.services.sheets.v4.model.*;

import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

/**
 * created by yash.zanwar on 23/12/18
 */
public class MakeCharts extends GoogleSheetApi {
    List<GridRange> gridRangeList = new ArrayList<GridRange>();
    GridRange tableRange = new GridRange();

    public void makeChart(String spreadsheetId, String title) throws IOException {
        Sheets service = getSheetsService();
        List<Request> requests = new ArrayList<Request>();

        BasicChartSpec basicChartSpec = new BasicChartSpec()
                .setChartType("LINE")
                .setLegendPosition("BOTTOM_LEGEND")
                .setAxis(setChartAxis("Model Numbers", "Sales"))
                .setDomains(setDomain())
                .setSeries(setSeries(0, 0, 8, 0, 4))
                .setHeaderCount(1);
        AddChartRequest addChartRequest = new AddChartRequest()
                .setChart(new EmbeddedChart()
                        .setSpec(new ChartSpec()
                                .setTitle(title)
                                .setBasicChart(basicChartSpec))
                        .setPosition(new EmbeddedObjectPosition()
                                .setOverlayPosition(new OverlayPosition()
                                        .setAnchorCell(new GridCoordinate()
                                                .setSheetId(0)
                                                .setRowIndex(10)
                                                .setColumnIndex(0)))));
        requests.add(new Request()
                .setAddChart(addChartRequest));

        BatchUpdateSpreadsheetRequest body = new BatchUpdateSpreadsheetRequest().setRequests(requests);
        BatchUpdateSpreadsheetResponse response = service.spreadsheets().batchUpdate(spreadsheetId, body).execute();
        System.out.printf("%d cells updated.", response.getReplies().size());
    }

    public List<BasicChartAxis> setChartAxis(String xAxisTitle, String yAxisTitle) {
        List<BasicChartAxis> listBasicChartAxis = new ArrayList<BasicChartAxis>();
        BasicChartAxis basicChartAxis = new BasicChartAxis().setPosition("BOTTOM_AXIS").setTitle(xAxisTitle);
        listBasicChartAxis.add(basicChartAxis);
        basicChartAxis = new BasicChartAxis().setPosition("LEFT_AXIS").setTitle(yAxisTitle);
        listBasicChartAxis.add(basicChartAxis);
        return listBasicChartAxis;
    }

    public List<BasicChartDomain> setDomain() {
        List<BasicChartDomain> basicChartDomainList = new ArrayList<BasicChartDomain>();
        GridRange tableRange = setRangeOfCells(0, 0, 7, 0, 1);
        gridRangeList.add(tableRange);
        BasicChartDomain basicChartDomain = new BasicChartDomain()
                .setDomain(new ChartData()
                        .setSourceRange(new ChartSourceRange()
                                .setSources(gridRangeList)));
        basicChartDomainList.add(basicChartDomain);
        return basicChartDomainList;
    }

    public List<BasicChartSeries> setSeries(int sheetId, int StartRow, int EndRow, int StartColumn, int noOfLines) {
        List<BasicChartSeries> basicChartSeriesList = new ArrayList<BasicChartSeries>();
        for (int i = 0; i < noOfLines; i++) {
            gridRangeList = new ArrayList<GridRange>();
            tableRange = setRangeOfCells(sheetId, StartRow, EndRow, StartColumn + i + 1, StartColumn + i + 2);
            gridRangeList.add(tableRange);
            BasicChartSeries basicChartSeries = new BasicChartSeries()
                    .setSeries(new ChartData()
                            .setSourceRange(new ChartSourceRange()
                                    .setSources(gridRangeList)))
                    .setTargetAxis("LEFT_AXIS");
            basicChartSeriesList.add(basicChartSeries);
        }
        return basicChartSeriesList;
    }
}
