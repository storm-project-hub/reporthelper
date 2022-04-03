package entity;

import annotation.ReportKey;

public class Employee {

    @ReportKey(name = "nameKey")
    private final String name;

    public Employee(String name) {
        this.name = name;
    }

    public String getName() {
        return name;
    }
}
