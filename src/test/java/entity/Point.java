package entity;

import annotation.ReportKey;
import enums.DataType;

public class Point {

    @ReportKey(type = DataType.NUMERIC)
    private String x;

    @ReportKey(type = DataType.NUMERIC)
    private double y;

    public Point() {
    }

    public void setX(String x) {
        this.x = x;
    }

    public void setY(double y) {
        this.y = y;
    }

    public String getX() {
        return x;
    }

    public double getY() {
        return y;
    }
}
