package com.jxls.writer.demo.model;


import java.util.ArrayList;
import java.util.List;

/**
 * Sample Department bean to demostrate main excel export features
 * @author Leonid Vysochyn
 */
public class Department {
    private String name;
    private Employee chief;
    private List<Employee> staff = new ArrayList<Employee>();

    public Department(String name) {
        this.name = name;
    }

    public Department(String name, Employee chief, List<Employee> staff) {
        this.name = name;
        this.chief = chief;
        this.staff = staff;
    }

    public void addEmployee(Employee employee) {
        staff.add(employee);
    }

    public String getName() {
        return name;
    }

    public void setName(String name) {
        this.name = name;
    }

    public Employee getChief() {
        return chief;
    }

    public void setChief(Employee chief) {
        this.chief = chief;
    }

    public List getStaff() {
        return staff;
    }

    public void setStaff(List staff) {
        this.staff = staff;
    }

    @Override
    public String toString() {
        return "Department{" +
                "name='" + name + '\'' +
                ", chief=" + chief +
                ", staff.size=" + staff.size() +
                '}';
    }
}
