package entity;

import annotation.ReportKey;
import enums.KeyType;

import java.util.List;

public class Company {

    @ReportKey(name = "nameKey")
    private final String name;

    @ReportKey(keyType = KeyType.COMPLEX)
    List<Employee> employees;

    public Company(String name, List<Employee> employees) {
        this.name = name;
        this.employees = employees;
    }

    public String getName() {
        return name;
    }
}
