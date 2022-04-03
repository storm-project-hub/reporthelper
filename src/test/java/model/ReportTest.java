package model;

import entity.Data;
import entity.DataRow;
import entity.DataSet;
import exception.IncorrectTemplateException;
import exception.ReportKeyException;
import org.apache.poi.ss.usermodel.Picture;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.*;
import org.junit.After;
import org.junit.Before;
import org.junit.Test;

import java.io.File;
import java.io.IOException;
import java.io.InputStream;
import java.nio.file.Files;
import java.util.*;

import static org.junit.Assert.*;

public class ReportTest {

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
    public void createReportFileTest() throws IOException, IncorrectTemplateException, ReportKeyException {
        template = getClass().getResourceAsStream("/template/Template1.xlsx");
        report.createReport(template, file);
        assertTrue(Files.exists(file.toPath()));
    }

    @Test
    public void createTemplateFileTest() throws IOException {
        report.createTemplate(file.getPath());
        assertTrue(Files.exists(file.toPath()));
    }

    @Test
    public void addComplexSheetToTemplateTest() throws IOException {
        report.createTemplate(file.getPath());
        XSSFWorkbook wb = new XSSFWorkbook(file.getPath());
        XSSFSheet sheet = wb.getSheet("complex_DataSet_dataRows");
        assertNotNull(sheet);
        wb.close();
    }

    @Test
    public void addCustomNameComplexSheetToTemplateTest() throws IOException {
        report.createTemplate(file.getPath());
        XSSFWorkbook wb = new XSSFWorkbook(file.getPath());
        XSSFSheet sheet = wb.getSheet("customKey");
        assertNotNull(sheet);
        wb.close();
    }

    @Test
    public void addDataBaseSheetToTemplateTest() throws IOException {
        report.createTemplate(file.getPath());
        XSSFWorkbook wb = new XSSFWorkbook(file.getPath());
        XSSFSheet sheet = wb.getSheet("REPORT_KEYS");
        assertNotNull(sheet);
        wb.close();
    }

    @Test
    public void fillDataBaseSheetTest() throws IOException {
        report.createTemplate(file.getPath());
        XSSFWorkbook wb = new XSSFWorkbook(file.getPath());
        XSSFSheet sheet = wb.getSheet("REPORT_KEYS");
        boolean isFilledKeys = false;
        List<String> keyList = new ArrayList<>();
        for (int i = 1; i <= sheet.getLastRowNum(); i++) {
            keyList.add(sheet.getRow(i).getCell(0).getStringCellValue());
        }
        Set<String> keys = new HashSet<>(Arrays.asList(
                "key_DataSet_text",
                "key_DataSet_temporaryKey",
                "key_DataSet_regularKey",
                "key_DataSet_doubleNumber",
                "key_DataSet_doubleInfinity",
                "key_DataSet_doubleNaN",
                "key_DataSet_intNumber",
                "key_DataSet_longNumber",
                "key_DataSet_floatNumber",
                "key_DataSet_defaultDate",
                "key_DataSet_negativeLong",
                "key_DataSet_minValueLong",
                "key_DataSet_anotherFormatDate",
                "key_DataSet_defaultTime",
                "key_DataSet_anotherFormatTime",
                "key_DataSet_imagePath",
                "key_DataRow_imagePath",
                "key_DataRow_text",
                "key_Data_text",
                "key_Point_x",
                "key_Point_y",
                "key_counter",
                "complex_DataRow_dataList",
                "complex_DataSet_dataRows",
                "complex_Data_pointList",
                "customKey"
        ));
        if (keyList.size() == keys.size()) {
            isFilledKeys = keys.containsAll(keyList);
        }
        assertTrue(isFilledKeys);
        wb.close();
    }

    @Test
    public void deleteTempListTest() throws IOException, IncorrectTemplateException, ReportKeyException {
        template = getClass().getResourceAsStream("/template/Template1.xlsx");
        report.createReport(template, file);
        XSSFWorkbook wb = new XSSFWorkbook(file.getPath());
        boolean isNormalised = true;
        Iterator<Sheet> iterator = wb.sheetIterator();
        while (iterator.hasNext()) {
            String name = iterator.next().getSheetName();
            if (name.contains("complex_") || name.contains("REPORT_KEYS") || name.contains("customKey")) {
                isNormalised = false;
                break;
            }
        }
        wb.close();
        assertTrue(isNormalised);
    }

    @Test
    public void fillRegularKeyWhenDataIsNullTest() throws IOException, IncorrectTemplateException, ReportKeyException {
        template = getClass().getResourceAsStream("/template/Template2.xlsx");
        report.createReport(template, file);
        XSSFWorkbook wb = new XSSFWorkbook(file.getPath());
        String actual = wb.getSheetAt(0).getRow(0).getCell(0).getStringCellValue();
        wb.close();
        assertEquals("", actual);
    }

    @Test
    public void deleteTemporaryKeyWhenDataIsNullTest() throws IOException, IncorrectTemplateException, ReportKeyException {
        template = getClass().getResourceAsStream("/template/Template3.xlsx");
        report.createReport(template, file);
        XSSFWorkbook wb = new XSSFWorkbook(file.getPath());
        String actual = wb.getSheetAt(0).getRow(0).getCell(0).getStringCellValue();
        String actual2 = wb.getSheetAt(0).getRow(0).getCell(2).getStringCellValue();
        wb.close();
        assertEquals("newValue", actual);
        assertEquals("newValue", actual2);
    }

    @Test
    public void deleteTemporaryKeyToMergedRegionWhenDataIsNullTest() throws IOException, IncorrectTemplateException, ReportKeyException {
        template = getClass().getResourceAsStream("/template/Template21.xlsx");
        report.createReport(template, file);
        XSSFWorkbook wb = new XSSFWorkbook(file.getPath());
        String actual = wb.getSheetAt(0).getRow(0).getCell(0).getStringCellValue();
        String actual2 = wb.getSheetAt(0).getRow(0).getCell(3).getStringCellValue();
        int numRegions = wb.getSheetAt(0).getNumMergedRegions();
        wb.close();
        assertEquals("SomeText", actual2);
        assertEquals("", actual);
        assertEquals(0, numRegions);
    }

    @Test
    public void fillTemporaryKeyWithSecondKeyInRowTest() throws IOException, IncorrectTemplateException, ReportKeyException {
        template = getClass().getResourceAsStream("/template/Template20.xlsx");
        report.createReport(template, file);
        XSSFWorkbook wb = new XSSFWorkbook(file.getPath());
        String actual = wb.getSheetAt(0).getRow(0).getCell(0).getStringCellValue();
        String actual2 = wb.getSheetAt(0).getRow(0).getCell(2).getStringCellValue();
        wb.close();
        assertEquals("SomeText", actual2);
        assertEquals("", actual);
    }

    @Test
    public void deleteTemporaryKeyWithMergedRegionInRowTest() throws IOException, IncorrectTemplateException, ReportKeyException {
        template = getClass().getResourceAsStream("/template/Template22.xlsx");
        report.createReport(template, file);
        XSSFWorkbook wb = new XSSFWorkbook(file.getPath());
        String actual = wb.getSheetAt(0).getRow(1).getCell(0).getStringCellValue();
        String actual2 = wb.getSheetAt(0).getRow(1).getCell(7).getStringCellValue();
        String actual3 = wb.getSheetAt(0).getRow(0).getCell(3).getStringCellValue();
        wb.close();
        assertEquals("", actual);
        assertEquals("SomeValue", actual2);
        assertEquals("SomeText", actual3);
    }

    @Test
    public void copyCellStyleTest() throws IOException, IncorrectTemplateException, ReportKeyException {
        template = getClass().getResourceAsStream("/template/Template16.xlsx");
        report.createReport(template, file);
        XSSFWorkbook wb = new XSSFWorkbook(file.getPath());
        XSSFCellStyle actualCellStyle = wb.getSheetAt(0).getRow(0).getCell(0).getCellStyle();

        template = getClass().getResourceAsStream("/template/Template16.xlsx");
        XSSFWorkbook templateWb = new XSSFWorkbook(template);
        XSSFSheet sheet = templateWb.getSheet("complex_DataSet_dataRows");
        XSSFCellStyle expectedCellStyle = sheet.getRow(0).getCell(0).getCellStyle();

        wb.close();
        templateWb.close();
        assertEquals(expectedCellStyle, actualCellStyle);
    }

    @Test
    public void autoSizeImageByCellSizeTest() throws IOException, IncorrectTemplateException, ReportKeyException {
        template = getClass().getResourceAsStream("/template/Template17.xlsx");
        report.createReport(template, file);
        XSSFWorkbook wb = new XSSFWorkbook(file.getPath());
        XSSFSheet sheet = wb.getSheetAt(0);
        XSSFDrawing drawing = sheet.createDrawingPatriarch();
        double imageHeight = 0;
        double imageWidth = 0;
        for (XSSFShape shape : drawing.getShapes()) {
            if (shape instanceof Picture) {
                XSSFPicture picture = (XSSFPicture) shape;
                imageHeight = picture.getCTPicture().getSpPr().getXfrm().getExt().getCy() / 9525.0;
                imageWidth = picture.getCTPicture().getSpPr().getXfrm().getExt().getCx() / 9525.0;
            }
        }

        template = getClass().getResourceAsStream("/template/Template17.xlsx");
        XSSFWorkbook templateWb = new XSSFWorkbook(template);
        XSSFSheet sheetTemp = templateWb.getSheetAt(0);
        double cellWidth = sheetTemp.getColumnWidthInPixels(1);
        double cellHeight = sheetTemp.getRow(1).getHeight() / 15.0;

        boolean isSameWidth = (Math.abs(cellWidth - imageWidth) / cellWidth * 100) < 0.1;
        boolean isSameHeight = (Math.abs(cellHeight - imageHeight) / cellHeight * 100) < 0.1;
        boolean isAutoSized = (isSameWidth && imageHeight <= cellHeight) ||
                (isSameHeight && imageWidth <= cellWidth);
        assertTrue(isAutoSized);
        wb.close();
        templateWb.close();
    }

    @Test
    public void autoSizeImageByMergedRegionTest() throws IOException, IncorrectTemplateException, ReportKeyException {
        template = getClass().getResourceAsStream("/template/Template18.xlsx");
        report.createReport(template, file);

        XSSFWorkbook wb = new XSSFWorkbook(file.getPath());
        XSSFSheet sheet = wb.getSheetAt(0);
        XSSFDrawing drawing = sheet.createDrawingPatriarch();
        double imageHeight = 0;
        double imageWidth = 0;
        for (XSSFShape shape : drawing.getShapes()) {
            if (shape instanceof Picture) {
                XSSFPicture picture = (XSSFPicture) shape;
                imageHeight = picture.getCTPicture().getSpPr().getXfrm().getExt().getCy() / 9525.0;
                imageWidth = picture.getCTPicture().getSpPr().getXfrm().getExt().getCx() / 9525.0;
            }
        }

        template = getClass().getResourceAsStream("/template/Template18.xlsx");
        XSSFWorkbook templateWb = new XSSFWorkbook(template);
        XSSFSheet sheetTemp = templateWb.getSheetAt(0);
        double regionWidth = 0;
        double regionHeight = 0;
        CellRangeAddress mergedRegion = sheetTemp.getMergedRegions().get(0);
        if (mergedRegion != null) {
            for (int i = mergedRegion.getFirstColumn(); i <= mergedRegion.getLastColumn(); i++) {
                regionWidth += sheetTemp.getColumnWidthInPixels(i);
            }
            for (int i = mergedRegion.getFirstRow(); i <= mergedRegion.getLastRow(); i++) {
                regionHeight += sheetTemp.getRow(i).getHeight() / 15.0;
            }
        }
        boolean isSameWidth = (Math.abs(regionWidth - imageWidth) / regionWidth * 100) < 0.1;
        boolean isSameHeight = (Math.abs(regionHeight - imageHeight) / regionHeight * 100) < 0.1;
        boolean isAutoSized = (isSameWidth && imageHeight <= regionHeight) ||
                (isSameHeight && imageWidth <= regionWidth);
        assertTrue(isAutoSized);

        wb.close();
        templateWb.close();
    }

    @Test
    public void autoSizeImageByWidthAndNormalizeRowTest() throws IOException, IncorrectTemplateException, ReportKeyException {
        template = getClass().getResourceAsStream("/template/Template19.xlsx");
        report.createReport(template, file);
        XSSFWorkbook wb = new XSSFWorkbook(file.getPath());
        XSSFSheet sheet = wb.getSheetAt(0);
        XSSFDrawing drawing = sheet.createDrawingPatriarch();
        double imageHeight = 0;
        double imageWidth = 0;
        for (XSSFShape shape : drawing.getShapes()) {
            if (shape instanceof Picture) {
                XSSFPicture picture = (XSSFPicture) shape;
                imageHeight = picture.getCTPicture().getSpPr().getXfrm().getExt().getCy() / 9525.0;
                imageWidth = picture.getCTPicture().getSpPr().getXfrm().getExt().getCx() / 9525.0;
            }
        }

        double regionWidth = 0;
        double regionHeight = 0;
        CellRangeAddress mergedRegion = sheet.getMergedRegions().get(0);
        if (mergedRegion != null) {
            for (int i = mergedRegion.getFirstColumn(); i <= mergedRegion.getLastColumn(); i++) {
                regionWidth += sheet.getColumnWidthInPixels(i);
            }
            regionHeight = sheet.getRow(1).getHeight() / 15.0;
        }

        wb.close();
        boolean isSameWidth = (Math.abs(regionWidth - imageWidth) / regionWidth * 100) < 0.1;
        boolean isSameHeight = (Math.abs(regionHeight - imageHeight) / regionHeight * 100) < 0.1;
        boolean isAutoSized = (isSameWidth && isSameHeight);
        assertTrue(isAutoSized);
    }

    @Test
    public void fillComplexKeyTest() throws IncorrectTemplateException, IOException, ReportKeyException {
        template = getClass().getResourceAsStream("/template/Template32.xlsx");
        report.createReport(template, file);
        XSSFWorkbook wb = new XSSFWorkbook(file.getPath());
        String actual = wb.getSheetAt(0).getRow(0).getCell(0).getStringCellValue();
        String actual2 = wb.getSheetAt(0).getRow(1).getCell(0).getStringCellValue();
        wb.close();
        assertEquals("someText1", actual);
        assertEquals("someText2", actual2);
    }

    @Test
    public void fillEmptyListTest() throws IOException, ReportKeyException, IncorrectTemplateException {
        template = getClass().getResourceAsStream("/template/Template33.xlsx");
        DataSet dataSet = new DataSet();
        dataSet.setDataRows(new ArrayList<>());
        Report report = new Report(dataSet);
        report.createReport(template, file);
        XSSFWorkbook wb = new XSSFWorkbook(file.getPath());
        String actual = wb.getSheetAt(0).getRow(0).getCell(0).getStringCellValue();
        String actual2 = wb.getSheetAt(0).getRow(1).getCell(0).getStringCellValue();
        wb.close();
        assertEquals("", actual);
        assertEquals("someValue", actual2);
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
