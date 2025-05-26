package com.example.timesheet;

import java.util.List;
import java.util.Map;

public class TimesheetRequest {
    private String department;
    private String employee;
    private List<Map<String, String>> table;

    // Getters and setters
    public String getDepartment() {
        return department;
    }

    public void setDepartment(String department) {
        this.department = department;
    }

    public String getEmployee() {
        return employee;
    }

    public void setEmployee(String employee) {
        this.employee = employee;
    }

    public List<Map<String, String>> getTable() {
        return table;
    }

    public void setTable(List<Map<String, String>> table) {
        this.table = table;
    }
}