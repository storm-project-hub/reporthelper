package model;

import entity.*;
import exception.FormatReportKeyException;
import exception.IdenticalReportKeyException;
import exception.IncorrectTemplateException;
import exception.ReportKeyException;
import org.junit.After;
import org.junit.Before;
import org.junit.Test;

import java.io.File;
import java.io.IOException;
import java.io.InputStream;
import java.nio.file.Files;
import java.util.List;

import static org.junit.Assert.assertEquals;
import static org.junit.Assert.assertThrows;

public class ExceptionReportTest {

    private File file;
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
    public void createReportWithSameKeysTest() {
        Exception exception = assertThrows(IdenticalReportKeyException.class, () -> {
            Employee employee = new Employee("employeeName");
            Company company = new Company("companyName", List.of(employee));
            Report report = new Report(company);
        });
        String expectedMessage = "Annotated fields have the identical names of ReportKey";
        String actualMessage = exception.getMessage();

        assertEquals(expectedMessage, actualMessage);
    }

    @Test
    public void formatReportKeyExceptionTest() throws ReportKeyException {
        InputStream template = getClass().getResourceAsStream("/template/Template45.xlsx");
        Point point = new Point();
        point.setX("text");
        Report report = new Report(point);
        Exception exception = assertThrows(FormatReportKeyException.class, () -> report.createReport(template, file));
        String expectedMessage = "Field datatype does not match the ReportKey datatype.";
        String actualMessage = exception.getMessage();

        assertEquals(expectedMessage, actualMessage);
    }

    @Test
    public void setComplexReportKeyOnSimpleFieldTest() {
        Exception exception = assertThrows(ReportKeyException.class, () -> {
            Report report = new Report(new IncorrectAnnotation("", ""));
        });
        String expectedMessage = "The annotation ReportKey does not match the field";
        String actualMessage = exception.getMessage();

        assertEquals(expectedMessage, actualMessage);
    }

    @Test
    public void checkCounterKeyOutsideComplexTest() {
        String template = "/template/Template35.xlsx";
        String expectedMessage = "The counter key is located outside the complex sheets.";
        checkIncorrectTemplateException(template, expectedMessage);
    }

    @Test
    public void checkComplexKeyInSheetWithSameNameTest() {
        String template = "/template/Template23.xlsx";
        String expectedMessage = "The sheet contains a complex key that references the same sheet.";
        checkIncorrectTemplateException(template, expectedMessage);
    }

    @Test
    public void checkLoopedComplexKeysTest() {
        String template = "/template/Template24.xlsx";
        String expectedMessage = "Complex keys have looped references.";
        checkIncorrectTemplateException(template, expectedMessage);
    }

    @Test
    public void checkLoopedComplexKeysTest2() {
        String template = "/template/Template38.xlsx";
        String expectedMessage = "Complex keys have looped references.";
        checkIncorrectTemplateException(template, expectedMessage);
    }

    @Test
    public void checkLoopedComplexKeysTest3() {
        String template = "/template/Template39.xlsx";
        String expectedMessage = "Complex keys have looped references.";
        checkIncorrectTemplateException(template, expectedMessage);
    }

    @Test
    public void checkLoopedComplexKeysTest4() {
        String template = "/template/Template40.xlsx";
        String expectedMessage = "Complex keys have looped references.";
        checkIncorrectTemplateException(template, expectedMessage);
    }

    @Test
    public void checkLoopedComplexKeysTest5() {
        String template = "/template/Template41.xlsx";
        String expectedMessage = "Complex keys have looped references.";
        checkIncorrectTemplateException(template, expectedMessage);
    }

    @Test
    public void checkLoopedComplexKeysTest6() {
        String template = "/template/Template42.xlsx";
        String expectedMessage = "Complex keys have looped references.";
        checkIncorrectTemplateException(template, expectedMessage);
    }

    @Test
    public void checkComplexKeyWithUnavailableDataTest() {
        String template = "/template/Template25.xlsx";
        String expectedMessage = "Incorrect use of the complex key. There is no access to the data object.";
        checkIncorrectTemplateException(template, expectedMessage);
    }

    @Test
    public void checkSingleKeyWithUnavailableDataTest() {
        String template = "/template/Template36.xlsx";
        String expectedMessage = "Incorrect use of the single key. There is no access to the data object in this sheet.";
        checkIncorrectTemplateException(template, expectedMessage);
    }

    @Test
    public void checkSingleKeyWithUnavailableDataInComplexTest() {
        String template = "/template/Template37.xlsx";
        String expectedMessage = "Incorrect use of the single key. There is no access to the data object in this sheet.";
        checkIncorrectTemplateException(template, expectedMessage);
    }

    private void checkIncorrectTemplateException(String path, String message) {
        Exception exception = assertThrows(IncorrectTemplateException.class, () -> {
            InputStream template = getClass().getResourceAsStream(path);
            report.createReport(template, file);
        });
        String actualMessage = exception.getMessage();
        assertEquals(message, actualMessage);
    }

    private DataSet createDataSet() {
        DataSet dataSet = new DataSet();
        dataSet.setText("text1");
        DataRow dataRow = new DataRow();
        dataRow.setText("text2");
        Data data = new Data();
        data.setPointList(List.of(new Point()));
        dataRow.setDataList(List.of(data));
        dataSet.setDataRows(List.of(dataRow));
        return dataSet;
    }
}
