package entity;

import annotation.ReportKey;
import enums.DataType;
import enums.KeyType;

import java.util.List;

public class DataRow {

    @ReportKey
    private String text;

    @ReportKey(type = DataType.IMAGE)
    private String imagePath;

    @ReportKey(keyType = KeyType.COMPLEX)
    private List<Data> dataList;

    public DataRow() {
    }

    public DataRow(String text, String imagePath, List<Data> dataList) {
        this.text = text;
        this.imagePath = imagePath;
        this.dataList = dataList;
    }

    public String getText() {
        return text;
    }

    public String getImagePath() {
        return imagePath;
    }

    public List<Data> getDataList() {
        return dataList;
    }

    public void setText(String text) {
        this.text = text;
    }

    public void setImagePath(String imagePath) {
        this.imagePath = imagePath;
    }

    public void setDataList(List<Data> dataList) {
        this.dataList = dataList;
    }
}
