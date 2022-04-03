package model;

import entity.Data;
import entity.DataRow;
import entity.DataSet;
import exception.IncorrectTemplateException;
import exception.ReportKeyException;
import org.apache.poi.ss.usermodel.Picture;
import org.apache.poi.xssf.usermodel.*;
import org.junit.After;
import org.junit.Before;
import org.junit.Test;

import java.io.File;
import java.io.IOException;
import java.io.InputStream;
import java.nio.file.Files;
import java.util.ArrayList;
import java.util.List;

import static org.junit.Assert.assertArrayEquals;
import static org.junit.Assert.assertEquals;

public class InsertDataReportTest {

    private File file;
    private InputStream template;
    private Report report;

    public final String TEMP = System.getProperty("user.home") + "/TESTS";

    @Before
    public void before() throws ReportKeyException {
        new File(TEMP).mkdir();
        file = new File(TEMP + "/test.xlsx");
        report = new Report(createDataSet());
    }

    @After
    public void after() throws IOException {
        if (Files.exists(file.toPath())) Files.delete(file.toPath());
    }

    @Test
    public void insertTextTest() throws IOException, IncorrectTemplateException, ReportKeyException {
        template = getClass().getResourceAsStream("/template/Template1.xlsx");
        report.createReport(template, file);
        XSSFWorkbook wb = new XSSFWorkbook(file.getPath());
        String actual = wb.getSheetAt(0).getRow(0).getCell(0).getStringCellValue();
        wb.close();
        assertEquals("SomeText", actual);
    }

    @Test
    public void insertTextToMergedRegionTest() throws IOException, IncorrectTemplateException, ReportKeyException {
        template = getClass().getResourceAsStream("/template/Template5.xlsx");
        report.createReport(template, file);
        XSSFWorkbook wb = new XSSFWorkbook(file.getPath());
        String actual = wb.getSheetAt(0).getRow(1).getCell(1).getStringCellValue();
        wb.close();
        assertEquals("SomeText", actual);
    }

    @Test
    public void insertImageTest() throws IOException, IncorrectTemplateException, ReportKeyException {
        template = getClass().getResourceAsStream("/template/Template6.xlsx");
        report.createReport(template, file);
        XSSFWorkbook wb = new XSSFWorkbook(file.getPath());
        XSSFSheet sheet = wb.getSheetAt(0);
        XSSFDrawing drawing = sheet.createDrawingPatriarch();

        int columnIndex = 0;
        int rowIndex = 0;
        for (XSSFShape shape : drawing.getShapes()) {
            if (shape instanceof Picture) {
                XSSFPicture picture = (XSSFPicture) shape;
                columnIndex = picture.getClientAnchor().getCol1();
                rowIndex = picture.getClientAnchor().getRow1();
            }
        }
        wb.close();
        assertArrayEquals(new int[]{2, 1}, new int[]{columnIndex, rowIndex});
    }

    @Test
    public void insertImageToMergedRegionTest() throws IOException, IncorrectTemplateException, ReportKeyException {
        template = getClass().getResourceAsStream("/template/Template7.xlsx");
        report.createReport(template, file);
        XSSFWorkbook wb = new XSSFWorkbook(file.getPath());
        XSSFSheet sheet = wb.getSheetAt(0);
        XSSFDrawing drawing = sheet.createDrawingPatriarch();

        int columnIndex = 0;
        int rowIndex = 0;
        for (XSSFShape shape : drawing.getShapes()) {
            if (shape instanceof Picture) {
                XSSFPicture picture = (XSSFPicture) shape;
                columnIndex = picture.getClientAnchor().getCol1();
                rowIndex = picture.getClientAnchor().getRow1();
            }
        }
        wb.close();
        assertArrayEquals(new int[]{2, 1}, new int[]{columnIndex, rowIndex});
    }

    @Test
    public void insertDateByDefaultFormatInTextCellTypeTest() throws IOException, IncorrectTemplateException, ReportKeyException {
        template = getClass().getResourceAsStream("/template/Template8.xlsx");
        report.createReport(template, file);
        XSSFWorkbook wb = new XSSFWorkbook(file.getPath());
        String actual = wb.getSheetAt(0).getRow(0).getCell(0).getStringCellValue();
        wb.close();
        assertEquals("15.02.2022", actual);
    }

    @Test
    public void insertDateByAnotherFormatInTextCellTypeTest() throws IOException, IncorrectTemplateException, ReportKeyException {
        template = getClass().getResourceAsStream("/template/Template9.xlsx");
        report.createReport(template, file);
        XSSFWorkbook wb = new XSSFWorkbook(file.getPath());
        String actual = wb.getSheetAt(0).getRow(0).getCell(0).getStringCellValue();
        wb.close();
        assertEquals("2022-02-15", actual);
    }

    @Test
    public void insertDateToDateCellFormatTest() throws IOException, IncorrectTemplateException, ReportKeyException {
        template = getClass().getResourceAsStream("/template/Template10.xlsx");
        report.createReport(template, file);
        XSSFWorkbook wb = new XSSFWorkbook(file.getPath());
        long time = wb.getSheetAt(0).getRow(0).getCell(0).getDateCellValue().getTime();
        wb.close();
        assertEquals(1644924015000L, time);
    }

    @Test
    public void insertNegativeLongToDateCellFormatTest() throws IOException, IncorrectTemplateException, ReportKeyException {
        template = getClass().getResourceAsStream("/template/Template30.xlsx");
        report.createReport(template, file);
        XSSFWorkbook wb = new XSSFWorkbook(file.getPath());
        long time = wb.getSheetAt(0).getRow(0).getCell(0).getDateCellValue().getTime();
        wb.close();
        assertEquals(-2209000820000L, time);
    }

    @Test
    public void insertMinValueLongToDateCellFormatTest() throws IOException, IncorrectTemplateException, ReportKeyException {
        template = getClass().getResourceAsStream("/template/Template31.xlsx");
        report.createReport(template, file);
        XSSFWorkbook wb = new XSSFWorkbook(file.getPath());
        double time = wb.getSheetAt(0).getRow(0).getCell(0).getNumericCellValue();
        wb.close();
        assertEquals(-625795703.1339792, time, 0);
    }

    @Test
    public void insertTimeDefaultFormatInTextCellTypeTest() throws IOException, IncorrectTemplateException, ReportKeyException {
        template = getClass().getResourceAsStream("/template/Template11.xlsx");
        report.createReport(template, file);
        XSSFWorkbook wb = new XSSFWorkbook(file.getPath());
        String actual = wb.getSheetAt(0).getRow(0).getCell(0).getStringCellValue();
        wb.close();
        assertEquals("15:20:15", actual);
    }

    @Test
    public void insertTimeAnotherFormatInTextCellTypeTest() throws IOException, IncorrectTemplateException, ReportKeyException {
        template = getClass().getResourceAsStream("/template/Template12.xlsx");
        report.createReport(template, file);
        XSSFWorkbook wb = new XSSFWorkbook(file.getPath());
        String actual = wb.getSheetAt(0).getRow(0).getCell(0).getStringCellValue();
        wb.close();
        assertEquals("3:20:15 PM", actual);
    }

    @Test
    public void insertTimeToTimeCellFormatTest() throws IOException, IncorrectTemplateException, ReportKeyException {
        template = getClass().getResourceAsStream("/template/Template12-1.xlsx");
        report.createReport(template, file);
        XSSFWorkbook wb = new XSSFWorkbook(file.getPath());
        long time = wb.getSheetAt(0).getRow(0).getCell(0).getDateCellValue().getTime();
        wb.close();
        assertEquals(1644924015000L, time);
    }

    @Test
    public void insertIntegerTest() throws IOException, IncorrectTemplateException, ReportKeyException {
        template = getClass().getResourceAsStream("/template/Template13.xlsx");
        report.createReport(template, file);
        XSSFWorkbook wb = new XSSFWorkbook(file.getPath());
        double actual = wb.getSheetAt(0).getRow(0).getCell(0).getNumericCellValue();
        wb.close();
        assertEquals(10, actual, 0);
    }

    @Test
    public void insertLongTest() throws IOException, IncorrectTemplateException, ReportKeyException {
        template = getClass().getResourceAsStream("/template/Template43.xlsx");
        report.createReport(template, file);
        XSSFWorkbook wb = new XSSFWorkbook(file.getPath());
        double actual = wb.getSheetAt(0).getRow(0).getCell(0).getNumericCellValue();
        wb.close();
        assertEquals(2147483649.0, actual, 0);
    }

    @Test
    public void insertFloatTest() throws IOException, IncorrectTemplateException, ReportKeyException {
        template = getClass().getResourceAsStream("/template/Template44.xlsx");
        report.createReport(template, file);
        XSSFWorkbook wb = new XSSFWorkbook(file.getPath());
        double actual = wb.getSheetAt(0).getRow(0).getCell(0).getNumericCellValue();
        wb.close();
        assertEquals(10.5, actual, 0);
    }

    @Test
    public void insertDoubleToNumericCellTypeTest() throws IOException, IncorrectTemplateException, ReportKeyException {
        template = getClass().getResourceAsStream("/template/Template14.xlsx");
        report.createReport(template, file);
        XSSFWorkbook wb = new XSSFWorkbook(file.getPath());
        double actual = wb.getSheetAt(0).getRow(0).getCell(0).getNumericCellValue();
        wb.close();
        assertEquals(10.5, actual, 0);
    }

    @Test
    public void insertDoubleToTextCellTypeTest() throws IOException, IncorrectTemplateException, ReportKeyException {
        template = getClass().getResourceAsStream("/template/Template27.xlsx");
        report.createReport(template, file);
        XSSFWorkbook wb = new XSSFWorkbook(file.getPath());
        double actual = wb.getSheetAt(0).getRow(0).getCell(0).getNumericCellValue();
        wb.close();
        assertEquals(10.5, actual, 0);
    }

    @Test
    public void insertNegativeDoubleTest() throws IOException, IncorrectTemplateException, ReportKeyException {
        template = getClass().getResourceAsStream("/template/Template27.xlsx");
        DataSet dataSet = new DataSet();
        dataSet.setDoubleNumber(-10.5);
        Report report1 = new Report(dataSet);
        report1.createReport(template, file);
        XSSFWorkbook wb = new XSSFWorkbook(file.getPath());
        double actual = wb.getSheetAt(0).getRow(0).getCell(0).getNumericCellValue();
        wb.close();
        assertEquals(-10.5, actual, 0);
    }

    @Test
    public void insertDoubleNanToTextCellTypeTest() throws IOException, IncorrectTemplateException, ReportKeyException {
        template = getClass().getResourceAsStream("/template/Template26.xlsx");
        report.createReport(template, file);
        XSSFWorkbook wb = new XSSFWorkbook(file.getPath());
        String actual = wb.getSheetAt(0).getRow(0).getCell(0).getErrorCellString();
        wb.close();
        assertEquals("#NUM!", actual);
    }

    @Test
    public void insertDoubleNanToNumericCellTypeTest() throws IOException, IncorrectTemplateException, ReportKeyException {
        template = getClass().getResourceAsStream("/template/Template28.xlsx");
        report.createReport(template, file);
        XSSFWorkbook wb = new XSSFWorkbook(file.getPath());
        String actual = wb.getSheetAt(0).getRow(0).getCell(0).getErrorCellString();
        wb.close();
        assertEquals("#NUM!", actual);
    }

    @Test
    public void insertDoubleInfinityToNumericCellTypeTest() throws IOException, IncorrectTemplateException, ReportKeyException {
        template = getClass().getResourceAsStream("/template/Template29.xlsx");
        report.createReport(template, file);
        XSSFWorkbook wb = new XSSFWorkbook(file.getPath());
        String actual = wb.getSheetAt(0).getRow(0).getCell(0).getErrorCellString();
        wb.close();
        assertEquals("#DIV/0!", actual);
    }

    @Test
    public void insertCounterKeyWhenMoreOneElementTest() throws IOException, IncorrectTemplateException, ReportKeyException {
        template = getClass().getResourceAsStream("/template/Template34.xlsx");
        report.createReport(template, file);
        XSSFWorkbook wb = new XSSFWorkbook(file.getPath());
        XSSFSheet sheet = wb.getSheetAt(0);
        double actual = sheet.getRow(0).getCell(0).getNumericCellValue();
        double actual2 = sheet.getRow(1).getCell(0).getNumericCellValue();
        String actual3 = sheet.getRow(2).getCell(0).getStringCellValue();
        assertEquals(1.0, actual, 0);
        assertEquals(2.0, actual2, 0);
        assertEquals("someValue", actual3);
        wb.close();
    }

    @Test
    public void insertCounterKeyWhenOneElementTest() throws IOException, ReportKeyException, IncorrectTemplateException {
        template = getClass().getResourceAsStream("/template/Template34.xlsx");
        DataSet dataSet = new DataSet();
        dataSet.setDataRows(List.of(new DataRow()));
        Report report = new Report(dataSet);
        report.createReport(template, file);
        XSSFWorkbook wb = new XSSFWorkbook(file.getPath());
        XSSFSheet sheet = wb.getSheetAt(0);
        double actual = sheet.getRow(0).getCell(0).getNumericCellValue();
        String actual2 = sheet.getRow(1).getCell(0).getStringCellValue();
        assertEquals(1.0, actual, 0);
        assertEquals("someValue", actual2);
        wb.close();
    }

    @Test
    public void insertCounterKeyWhenEmptyListTest() throws IOException, ReportKeyException, IncorrectTemplateException {
        template = getClass().getResourceAsStream("/template/Template34.xlsx");
        DataSet dataSet = new DataSet();
        dataSet.setDataRows(new ArrayList<>());
        Report report = new Report(dataSet);
        report.createReport(template, file);
        XSSFWorkbook wb = new XSSFWorkbook(file.getPath());
        XSSFSheet sheet = wb.getSheetAt(0);
        String actual = sheet.getRow(0).getCell(0).getStringCellValue();
        String actual2 = sheet.getRow(1).getCell(0).getStringCellValue();
        assertEquals("", actual);
        assertEquals("someValue", actual2);
        wb.close();
    }

    @Test
    public void insertCounterKeyWhenListIsNullTest() throws IOException, ReportKeyException, IncorrectTemplateException {
        template = getClass().getResourceAsStream("/template/Template34.xlsx");
        DataSet dataSet = new DataSet();
        Report report = new Report(dataSet);
        report.createReport(template, file);
        XSSFWorkbook wb = new XSSFWorkbook(file.getPath());
        XSSFSheet sheet = wb.getSheetAt(0);
        String actual = sheet.getRow(0).getCell(0).getStringCellValue();
        String actual2 = sheet.getRow(1).getCell(0).getStringCellValue();
        assertEquals("", actual);
        assertEquals("someValue", actual2);
        wb.close();
    }

    private DataSet createDataSet() {
        List<Data> dataList = new ArrayList<>();
        dataList.add(new Data("Data1"));
        dataList.add(new Data("Data2"));
        List<DataRow> dataRows = new ArrayList<>();
        dataRows.add(new DataRow("someText1", "src/test/resources/img/test.jpg", dataList));
        dataRows.add(new DataRow("someText2", "src/test/resources/img/test.jpg", dataList));

        return new DataSet(
                "SomeText",
                10.5,
                10,
                1644924015000L,
                "src/test/resources/img/test.jpg",
                dataRows);
    }
}
