package entity;

import annotation.ReportKey;
import enums.KeyType;

import java.util.List;

public class Data {

    @ReportKey
    private String text;

    @ReportKey(keyType = KeyType.COMPLEX)
    private List<Point> pointList;

    public Data() {
    }

    public Data(String text) {
        this.text = text;
    }

    public String getText() {
        return text;
    }

    public void setPointList(List<Point> pointList) {
        this.pointList = pointList;
    }

    public List<Point> getPointList() {
        return pointList;
    }
}
