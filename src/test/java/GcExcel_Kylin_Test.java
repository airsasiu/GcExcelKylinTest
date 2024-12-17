import com.grapecity.documents.excel.*;
import com.grapecity.documents.excel.drawing.ChartType;
import com.grapecity.documents.excel.drawing.IShape;
import com.grapecity.documents.excel.drawing.RowCol;
import com.grapecity.documents.excel.template.DataSource.JsonDataSource;
import org.junit.Test;

import java.io.File;
import java.nio.file.Files;
import java.util.GregorianCalendar;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertTrue;

public class GcExcel_Kylin_Test {

    Object SOURCEDATA = new Object[][]{
            {"Name", "City", "Birthday", "Eye color", "Weight", "Height"},
            {"Richard", "New York", new GregorianCalendar(1968, 5, 8), "Blue", 67, 165},
            {"Nia", "New York", new GregorianCalendar(1972, 6, 3), "Brown", 62, 134},
            {"Jared", "New York", new GregorianCalendar(1964, 2, 2), "Hazel", 72, 180},
            {"Natalie", "Washington", new GregorianCalendar(1972, 7, 8), "Blue", 66, 163},
            {"Damon", "Washington", new GregorianCalendar(1986, 1, 2), "Hazel", 76, 176},
            {"Angela", "Washington", new GregorianCalendar(1993, 1, 15), "Brown", 68, 145}
    };

    Object PIVOTSOURCEDATA = new Object[][]{
            {"Order ID", "Product", "Category", "Amount", "Date", "Country"},
            {1, "Bose 785593-0050", "Consumer Electronics", 4270, new GregorianCalendar(2018, 0, 6), "United States"},
            {2, "Canon EOS 1500D", "Consumer Electronics", 8239, new GregorianCalendar(2018, 0, 7), "United Kingdom"},
            {3, "Haier 394L 4Star", "Consumer Electronics", 617, new GregorianCalendar(2018, 0, 8), "United States"},
            {4, "IFB 6.5 Kg FullyAuto", "Consumer Electronics", 8384, new GregorianCalendar(2018, 0, 10), "Canada"},
            {5, "Mi LED 40inch", "Consumer Electronics", 2626, new GregorianCalendar(2018, 0, 10), "Germany"},
            {6, "Sennheiser HD 4.40-BT", "Consumer Electronics", 3610, new GregorianCalendar(2018, 0, 11), "United States"},
            {7, "Iphone XR", "Mobile", 9062, new GregorianCalendar(2018, 0, 11), "Australia"},
            {8, "OnePlus 7Pro", "Mobile", 6906, new GregorianCalendar(2018, 0, 16), "New Zealand"},
            {9, "Redmi 7", "Mobile", 2417, new GregorianCalendar(2018, 0, 16), "France"},
            {10, "Samsung S9", "Mobile", 7431, new GregorianCalendar(2018, 0, 16), "Canada"},
            {11, "OnePlus 7Pro", "Mobile", 8250, new GregorianCalendar(2018, 0, 16), "Germany"},
            {12, "Redmi 7", "Mobile", 7012, new GregorianCalendar(2018, 0, 18), "United States"},
            {13, "Bose 785593-0050", "Consumer Electronics", 1903, new GregorianCalendar(2018, 0, 20), "Germany"},
            {14, "Canon EOS 1500D", "Consumer Electronics", 2824, new GregorianCalendar(2018, 0, 22), "Canada"},
            {15, "Haier 394L 4Star", "Consumer Electronics", 6946, new GregorianCalendar(2018, 0, 24), "France"},
    };

    // manage worksheets
    @Test
    public void GcExcel_Test_01() {
        Workbook wb = new Workbook();
        wb.getWorksheets().add();
        wb.getWorksheets().addBefore(wb.getWorksheets().get(0));
        wb.getWorksheets().addAfter(wb.getWorksheets().get(1));
        wb.getWorksheets().get(2).setName("Product Plan");

        assertEquals(wb.getWorksheets().getCount(), 4);
        assertEquals(wb.getWorksheets().get(0).getName(), "Sheet3");
        assertEquals(wb.getWorksheets().get(1).getName(), "Sheet1");
        assertEquals(wb.getWorksheets().get(2).getName(), "Product Plan");
        assertEquals(wb.getWorksheets().get(3).getName(), "Sheet2");
    }

    // set and get value
    @Test
    public void GcExcel_Test_02() {
        Workbook wb = new Workbook();
        IWorksheet sheet = wb.getWorksheets().get(0);

        sheet.getRange("A1:F7").setValue(SOURCEDATA);

        assertEquals(sheet.getRange("A1").getValue(), "Name");
        assertEquals(sheet.getRange("B2").getValue(), "New York");
        assertEquals(sheet.getRange("C3").getValue().toString(), "1972-07-03T00:00");
        assertEquals(sheet.getRange("D4").getValue(), "Hazel");
        assertEquals(sheet.getRange("E5").getValue(), 66.0);
        assertEquals(sheet.getRange("F6").getValue(), 176.0);
    }

    // set and get style
    @Test
    public void GcExcel_Test_03() {
        Workbook wb = new Workbook();
        IWorksheet sheet = wb.getActiveSheet();
        sheet.getRange("A1").setValue("blue content");
        sheet.getRange("A1").getInterior().setColor(Color.GetGreen());
        sheet.getRange("A1").getInterior().setColor(Color.GetBlue());

        assertEquals(sheet.getRange("A1").getValue(), "blue content");
        assertEquals(sheet.getRange("A1").getInterior().getColor(), Color.GetBlue());
    }

    // change worksheet layout
    @Test
    public void GcExcel_Test_04() {
        Workbook wb = new Workbook();
        IWorksheet sheet = wb.getActiveSheet();

        sheet.getRange("1:2").setRowHeight(50);
        sheet.getRange("C:D").setColumnWidth(20);

        assertEquals(sheet.getRange("A1").getRowHeight(), 50);
        assertEquals(sheet.getRange("A2").getRowHeight(), 50);
        assertEquals(sheet.getRange("A1").getColumnWidth(), 8.42578125);
        assertEquals(sheet.getRange("A2").getColumnWidth(), 8.42578125);

        assertEquals(sheet.getRange("C1").getRowHeight(), 50);
        assertEquals(sheet.getRange("C2").getRowHeight(), 50);
        assertEquals(sheet.getRange("C1").getColumnWidth(), 20);
        assertEquals(sheet.getRange("C2").getColumnWidth(), 20);
    }

    // formula
    @Test
    public void GcExcel_Test_05() {
        Workbook wb = new Workbook();
        wb.getOptions().getFormulas().setMaximumIterations(1000);
        wb.getOptions().getFormulas().setMaximumChange(0.000001);

        IWorksheet sheet = wb.getActiveSheet();
        sheet.getRange("A1:A4").setValue(new String[]{"Loan Amount", "Term in Months", "Interest Rate", "Payment"});

        sheet.getRange("B1").setValue(100000);
        sheet.getRange("B1").setNumberFormat("$#,##0");

        sheet.getRange("B2").setValue(180);

        sheet.getRange("B3").setNumberFormat("0.00%");
        sheet.getRange("B4").setFormula("=PMT(B3/12,B2,B1)");
        sheet.getRange("B4").setNumberFormat("$#,##0");
        sheet.getRange("B4").goalSeek(-900, sheet.getRange("B3"));

        assertEquals(sheet.getRange("B1").getValue(), 100000.0);
        assertEquals(sheet.getRange("B2").getValue(), 180.0);
        assertEquals(sheet.getRange("B3").getValue(), 0.07020951110333773);
        assertEquals(sheet.getRange("B4").getValue(), -900.0000006108976);
    }

    // chart
    @Test
    public void GcExcel_Test_06() {
        Workbook wb = new Workbook();

        IWorksheet sheet = wb.getWorksheets().get(0);

        IShape shape = sheet.getShapes().addChart(ChartType.ColumnClustered, 300, 10, 300, 300);
        sheet.getRange("A1:D6").setValue(new Object[][]{
                {null, "S1", "S2", "S3"},
                {"Item1", 10, 25, 25},
                {"Item2", -51, -36, 27},
                {"Item3", 52, -85, -30},
                {"Item4", 22, 65, 65},
                {"Item5", 23, 69, 69}
        });

        shape.getChart().getSeriesCollection().add(sheet.getRange("A1:D6"), RowCol.Columns, true, true);

        assertEquals(sheet.getShapes().getCount(), 1);
        assertEquals(sheet.getShapes().get(0).getChart().getChartTitle().getText(), "");
        assertEquals(sheet.getShapes().get(0).getChart().getChartType(), ChartType.ColumnClustered);
        assertEquals(sheet.getShapes().get(0).getChart().getSeriesCollection().getCount(), 3);
        assertEquals(sheet.getShapes().get(0).getChart().getSeriesCollection().get(0).getName(), "S1");
        assertEquals(sheet.getShapes().get(0).getChart().getSeriesCollection().get(1).getName(), "S2");
        assertEquals(sheet.getShapes().get(0).getChart().getSeriesCollection().get(2).getName(), "S3");
    }

    // table
    @Test
    public void GcExcel_Test_07() {
        Workbook wb = new Workbook();
        IWorksheet sheet = wb.getActiveSheet();

        sheet.getRange("A1:F7").setValue(SOURCEDATA);
        sheet.getRange("A:F").setColumnWidth(15);

        sheet.getTables().add(sheet.getRange("A1:F7"), true);

        ITable table = sheet.getTables().get(0);
        table.setShowTotals(true);

        assertEquals(sheet.getRange("F8").getValue(), 963.0);
        assertEquals(table.getColumns().getCount(), 6);
        assertEquals(table.getRows().getCount(), 6);

        table.getRange().autoFilter(0, new Object[]{"Richard", "Nia"}, AutoFilterOperator.Values);

        assertEquals(sheet.getRange("F8").getValue(), 299.0);
        assertEquals(table.getColumns().getCount(), 6);
        assertEquals(table.getRows().getCount(), 6);
    }

    // pivotTable
    @Test
    public void GcExcel_Test_08(){
        Workbook wb = new Workbook();
        IWorksheet sheet = wb.getActiveSheet();
        sheet.getRange("G1:L16").setValue(PIVOTSOURCEDATA);
        sheet.getRange("G:L").setColumnWidth(15);

        IPivotCache pivotCache = wb.getPivotCaches().create(sheet.getRange("G1:L16"));
        IPivotTable pivotTable = sheet.getPivotTables().add(pivotCache, sheet.getRange("A1"), "pivottable1");
        sheet.getRange("J1:J16").setNumberFormat("$#,##0.00");

        IPivotField field_Category = pivotTable.getPivotFields().get("Category");
        field_Category.setOrientation(PivotFieldOrientation.ColumnField);

        IPivotField field_Product = pivotTable.getPivotFields().get("Product");
        field_Product.setOrientation(PivotFieldOrientation.RowField);

        IPivotField field_Amount = pivotTable.getPivotFields().get("Amount");
        field_Amount.setOrientation(PivotFieldOrientation.DataField);
        field_Amount.setNumberFormat("$#,##0.00");

        IPivotField field_Country = pivotTable.getPivotFields().get("Country");
        field_Country.setOrientation(PivotFieldOrientation.PageField);
        sheet.getPivotTables().get(0).refresh();

        sheet.getRange("A:D").getEntireColumn().autoFit();

        assertEquals(sheet.getRange("B15").getValue(), 39419.00);
        assertEquals(sheet.getRange("C15").getValue(), 41078.00);
        assertEquals(sheet.getRange("D15").getValue(), 80497.00);
    }

    // sort data
    @Test
    public void GcExcel_Test_09(){
        Workbook wb = new Workbook();
        IWorksheet sheet = wb.getActiveSheet();

        sheet.getRange("A1:F7").setValue(SOURCEDATA);
        sheet.getRange("A:F").setColumnWidth(15);

        assertEquals(sheet.getRange("F2").getValue(),165.0);
        assertEquals(sheet.getRange("F3").getValue(),134.0);
        assertEquals(sheet.getRange("F4").getValue(),180.0);
        assertEquals(sheet.getRange("F5").getValue(),163.0);
        assertEquals(sheet.getRange("F6").getValue(),176.0);
        assertEquals(sheet.getRange("F7").getValue(),145.0);

        sheet.getRange("A2:F7").sort(sheet.getRange("F2:F7"), SortOrder.Ascending, SortOrientation.Columns);

        assertEquals(sheet.getRange("F2").getValue(),134.0);
        assertEquals(sheet.getRange("F3").getValue(),145.0);
        assertEquals(sheet.getRange("F4").getValue(),163.0);
        assertEquals(sheet.getRange("F5").getValue(),165.0);
        assertEquals(sheet.getRange("F6").getValue(),176.0);
        assertEquals(sheet.getRange("F7").getValue(),180.0);
    }

    // search data
    @Test
    public void GcExcel_Test_10(){
        Workbook wb = new Workbook();
        IWorksheet sheet = wb.getActiveSheet();

        final String CorrectWord = "Macro";
        sheet.getRange("A1:D5").setValue(CorrectWord);

        final String MisspelledWord = "marco";
        sheet.getRange("A2,C3,D1").setValue(MisspelledWord);

        assertEquals(sheet.getRange("A2").getValue(), "marco");
        assertEquals(sheet.getRange("C3").getValue(), "marco");
        assertEquals(sheet.getRange("D1").getValue(), "marco");

        IRange searchRange = sheet.getRange("A1:D5");
        while(true){
            IRange misspelledCell = searchRange.find(MisspelledWord);
            if(misspelledCell != null){
                misspelledCell.setValue(CorrectWord);
            }
            else{
                break;
            }
        }

        assertEquals(sheet.getRange("A2").getValue(), "Macro");
        assertEquals(sheet.getRange("C3").getValue(), "Macro");
        assertEquals(sheet.getRange("D1").getValue(), "Macro");
    }

    // condition format
    @Test
    public void GcExcel_Test_11(){
        String jsonText = "[" +
                "{\"Area\": \"North America\",\"City\": \"Chicago\",\"Category\": \"Consumer Electronics\",\"Name\": \"Bose 785593-0050\",\"Revenue\": 92800},\n" +
                "{\"Area\": \"North America\",\"City\": \"New York\",\"Category\": \"Consumer Electronics\",\"Name\": \"Bose 785593-0050\",\"Revenue\": 92800},\n" +
                "{\"Area\": \"South America\",\"City\": \"Santiago\",\"Category\": \"Consumer Electronics\",\"Name\": \"Bose 785593-0050\",\"Revenue\": 19550}\n" +
                "]";
        Workbook wb = new Workbook();
        IWorksheet sheet = wb.getActiveSheet();
        sheet.setDataSource(new JsonDataSource(jsonText));

        assertEquals(sheet.getRange("A1").getValue(), "North America");
        assertEquals(sheet.getRange("B2").getValue(), "New York");
        assertEquals(sheet.getRange("C3").getValue(), "Consumer Electronics");
        assertEquals(sheet.getRange("D2").getValue(), "Bose 785593-0050");
        assertEquals(sheet.getRange("E1").getValue(), 92800.0);
    }

    // to json
    @Test
    public void GcExcel_Test_12(){
        Workbook wb = new Workbook();
        IWorksheet sheet = wb.getActiveSheet();
        sheet.getRange("A1:F7").setValue(SOURCEDATA);
        ITop10 top10 = sheet.getRange("F2:F7").getFormatConditions().addTop10();
        top10.setRank(3);
        top10.setNumberFormat("0.00");
        top10.getInterior().setColor(Color.FromArgb(91,155,213));
        String json = top10.toJson();

        assertEquals(json,"{\"ranges\":[{\"row\":1,\"col\":5,\"rowCount\":6,\"colCount\":1}],\"stopIfTrue\":false,\"priority\":1,\"style\":{\"backColor\":\"rgb(91,155,213)\",\"formatter\":\"0.00\"},\"ruleType\":5,\"type\":0,\"rank\":3}");;
    }

    // save-load sjs
    @Test
    public void GcExcel_Test_13(){
        Workbook wb = new Workbook();
        IWorksheet sheet = wb.getActiveSheet();
        sheet.getRange("A1").setValue("test");
        wb.save("saveload.sjs");

        File file = new File("saveload.sjs");
        assertTrue(file.exists());

        wb = new Workbook();
        wb.open("saveload.sjs");
        assertEquals(wb.getWorksheets().get("Sheet1").getRange("A1").getValue(), "test");

        if(file.exists()){
            file.delete();
        }
    }

    // save-load xlsx
    @Test
    public void GcExcel_Test_14(){
        Workbook wb = new Workbook();
        IWorksheet sheet = wb.getActiveSheet();
        sheet.getRange("A1").setValue("test");
        wb.save("saveload.xlsx");

        File file = new File("saveload.xlsx");
        assertTrue(file.exists());

        wb = new Workbook();
        wb.open("saveload.xlsx");
        assertEquals(wb.getWorksheets().get("Sheet1").getRange("A1").getValue(), "test");

        if(file.exists()){
            file.delete();
        }
    }

    // export pdf
    @Test
    public void GcExcel_Test_15(){
        Workbook wb = new Workbook();
        IWorksheet sheet = wb.getActiveSheet();
        sheet.getRange("A1").setValue("test");
        wb.save("saveload.pdf");

        File file = new File("saveload.pdf");
        assertTrue(file.exists());

        if(file.exists()){
            file.delete();
        }
    }
}
