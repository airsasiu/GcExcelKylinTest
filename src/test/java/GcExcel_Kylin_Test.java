import com.grapecity.documents.excel.*;
import static org.junit.jupiter.api.Assertions.*;
import org.junit.Test;

import java.util.GregorianCalendar;

public class GcExcel_Kylin_Test {

    @Test
    public void GcExcel_Test_01(){
        Workbook wb = new Workbook();
        wb.getWorksheets().add();
        wb.getWorksheets().addBefore(wb.getWorksheets().get(0));
        wb.getWorksheets().addAfter(wb.getWorksheets().get(1));
        wb.getWorksheets().get(2).setName("Product Plan");

        assertEquals(wb.getWorksheets().getCount(),4);
        assertEquals(wb.getWorksheets().get(0).getName(),"Sheet3");
        assertEquals(wb.getWorksheets().get(1).getName(),"Sheet1");
        assertEquals(wb.getWorksheets().get(2).getName(),"Product Plan");
        assertEquals(wb.getWorksheets().get(3).getName(),"Sheet2");
    }

    @Test
    public void GcExcel_Test_02(){
        Workbook wb = new Workbook();
        IWorksheet sheet = wb.getWorksheets().get(0);
        Object data = new Object[][]{
                {"Name", "City", "Birthday", "Eye color", "Weight", "Height"},
                {"Richard", "New York", new GregorianCalendar(1968, 5, 8), "Blue", 67, 165},
                {"Nia", "New York", new GregorianCalendar(1972, 6, 3), "Brown", 62, 134},
                {"Jared", "New York", new GregorianCalendar(1964, 2, 2), "Hazel", 72, 180},
                {"Natalie", "Washington", new GregorianCalendar(1972, 7, 8), "Blue", 66, 163},
                {"Damon", "Washington", new GregorianCalendar(1986, 1, 2), "Hazel", 76, 176},
                {"Angela", "Washington", new GregorianCalendar(1993, 1, 15), "Brown", 68, 145}
        };
        sheet.getRange("A1:F7").setValue(data);

        assertEquals(sheet.getRange("A1").getValue(), "Name");
        assertEquals(sheet.getRange("B2").getValue(), "New York");
        assertEquals(sheet.getRange("C3").getValue().toString(), "1972-07-03T00:00");
        assertEquals(sheet.getRange("D4").getValue(), "Hazel");
        assertEquals(sheet.getRange("E5").getValue(), 66.0);
        assertEquals(sheet.getRange("F6").getValue(), 176.0);

    }
}
