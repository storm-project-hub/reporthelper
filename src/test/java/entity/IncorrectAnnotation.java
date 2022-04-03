package entity;

import annotation.ReportKey;
import enums.KeyType;

public class IncorrectAnnotation {

    @ReportKey
    private final String text;

    @ReportKey(keyType = KeyType.COMPLEX)
    private String list;

    public IncorrectAnnotation(String text, String list) {
        this.text = text;
        this.list = list;
    }

    public String getText() {
        return text;
    }

    public String getList() {
        return list;
    }

    public void setList(String list) {
        this.list = list;
    }
}
