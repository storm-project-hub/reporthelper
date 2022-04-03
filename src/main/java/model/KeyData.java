package model;

import annotation.ReportKey;

import java.lang.reflect.Field;

public class KeyData {

    private final String name;
    private final String description;
    private final ReportKey reportKey;
    private final Field field;

    public KeyData(String name, String description, ReportKey reportKey, Field field) {
        this.name = name;
        this.description = description;
        this.reportKey = reportKey;
        this.field = field;
    }

    public ReportKey getReportKey() {
        return reportKey;
    }

    public Field getField() {
        return field;
    }

    public String getName() {
        return name;
    }

    public String getDescription() {
        return description;
    }
}
