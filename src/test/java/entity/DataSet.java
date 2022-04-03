package entity;

import annotation.ReportKey;
import enums.DataType;
import enums.KeyType;

import java.util.ArrayList;
import java.util.List;

public class DataSet {

    @ReportKey
    private String text;

    @ReportKey(temporary = true)
    private String temporaryKey;

    @ReportKey
    private String regularKey;

    @ReportKey(type = DataType.NUMERIC)
    private double doubleNumber;

    @ReportKey(type = DataType.NUMERIC)
    private double doubleNaN;

    @ReportKey(type = DataType.NUMERIC)
    private double doubleInfinity;

    @ReportKey(type = DataType.NUMERIC)
    private int intNumber;

    @ReportKey(type = DataType.NUMERIC)
    private long longNumber;

    @ReportKey(type = DataType.NUMERIC)
    private float floatNumber;

    @ReportKey(type = DataType.DATE)
    private long defaultDate;

    @ReportKey(type = DataType.DATE)
    private long negativeLong;

    @ReportKey(type = DataType.DATE)
    private long minValueLong;

    @ReportKey(type = DataType.DATE, dateFormatPattern = "yyyy-MM-dd")
    private long anotherFormatDate;

    @ReportKey(type = DataType.TIME)
    private long defaultTime;

    @ReportKey(type = DataType.TIME, timeFormatPattern = "h:mm:ss aaa")
    private long anotherFormatTime;

    @ReportKey(type = DataType.IMAGE)
    private String imagePath;

    @ReportKey(keyType = KeyType.COMPLEX)
    private List<DataRow> dataRows;

    @ReportKey(name = "customKey", keyType = KeyType.COMPLEX)
    private List<DataRow> customNameList;

    public DataSet() {
    }

    public DataSet(String text, double doubleNumber,
                   int intNumber, long date, String imagePath, List<DataRow> dataRows) {
        this.text = text;
        this.temporaryKey = null;
        this.regularKey = null;
        this.doubleNumber = doubleNumber;
        this.doubleNaN = Double.NaN;
        this.doubleInfinity = Double.POSITIVE_INFINITY;
        this.intNumber = intNumber;
        this.longNumber = 2147483649L;
        this.floatNumber = 10.5f;
        this.defaultDate = date;
        this.negativeLong = -2209000820000L;
        this.minValueLong = Long.MIN_VALUE;
        this.anotherFormatDate = date;
        this.defaultTime = date;
        this.anotherFormatTime = date;
        this.imagePath = imagePath;
        this.dataRows = dataRows;
        this.customNameList = new ArrayList<>();
    }

    public void setText(String text) {
        this.text = text;
    }

    public void setTemporaryKey(String temporaryKey) {
        this.temporaryKey = temporaryKey;
    }

    public void setRegularKey(String regularKey) {
        this.regularKey = regularKey;
    }

    public void setDoubleNumber(double doubleNumber) {
        this.doubleNumber = doubleNumber;
    }

    public void setIntNumber(int intNumber) {
        this.intNumber = intNumber;
    }

    public void setDefaultDate(long defaultDate) {
        this.defaultDate = defaultDate;
    }

    public void setAnotherFormatDate(long anotherFormatDate) {
        this.anotherFormatDate = anotherFormatDate;
    }

    public void setDefaultTime(long defaultTime) {
        this.defaultTime = defaultTime;
    }

    public void setAnotherFormatTime(long anotherFormatTime) {
        this.anotherFormatTime = anotherFormatTime;
    }

    public void setImagePath(String imagePath) {
        this.imagePath = imagePath;
    }

    public void setDataRows(List<DataRow> dataRows) {
        this.dataRows = dataRows;
    }

    public List<DataRow> getCustomNameList() {
        return customNameList;
    }

    public void setCustomNameList(List<DataRow> customNameList) {
        this.customNameList = customNameList;
    }

    public void setDoubleNaN(double doubleNaN) {
        this.doubleNaN = doubleNaN;
    }

    public void setDoubleInfinity(double doubleInfinity) {
        this.doubleInfinity = doubleInfinity;
    }

    public void setNegativeLong(long negativeLong) {
        this.negativeLong = negativeLong;
    }

    public void setMinValueLong(long minValueLong) {
        this.minValueLong = minValueLong;
    }
}
