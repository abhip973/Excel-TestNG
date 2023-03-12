package abhiLearns.testNG_Excel_Integration;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.testng.annotations.DataProvider;
import org.testng.annotations.Test;

import java.io.FileInputStream;
import java.io.IOException;

public class DataProvide {

    @Test(dataProvider = "getData")
    public void testData(String greet, String id, String day) {

        System.out.println(greet + id + day);
    }


    @DataProvider(name = "getData")
    public Object[][] getData() throws IOException {

        DataFormatter formatter = new DataFormatter();
        FileInputStream file = new FileInputStream("src/main/java/Excel/DataSheet.xlsx");
        XSSFWorkbook workbook = new XSSFWorkbook(file);
        XSSFSheet sheet = workbook.getSheet("Sheet1");
        Row row;

        int total_rows = sheet.getPhysicalNumberOfRows();
        XSSFRow headerRow = sheet.getRow(0);
        int columnNum = headerRow.getLastCellNum();

        Object[][] data = new Object[total_rows - 1][columnNum];

        for (int i = 0; i < total_rows - 1; i++) {
            row = sheet.getRow(i + 1);
            for (int j = 0; j < columnNum; j++) {
                Cell cell = row.getCell(j);
                data[i][j] = formatter.formatCellValue(cell);
//                System.out.println(data[i][j]);
            }

        }

        return data;
    }
}
