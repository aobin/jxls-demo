package com.jxls.plus.demo.model;


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
    private String link;

    public Department(String name) {
        this.name = name;
    }

    public Department(String name, Employee chief, List<Employee> staff) {
        this.name = name;
        this.chief = chief;
        this.staff = staff;
    }

    public static List<Department> generate(int depCount, int employeeCount){
        List<Department> departments = new ArrayList<Department>();
        for(int index = 0; index < depCount; index++){
            Department dep = new Department("Dep " + index);
            dep.setChief( Employee.generateOne("ch" + index));
            dep.setStaff( Employee.generate(employeeCount) );
            departments.add( dep );
        }
        return departments;
    }

    public void addEmployee(Employee employee) {
        staff.add(employee);
    }

    public int getHeadcount() {
        return staff.size();
    }

    public String getName() {
        return name;
    }

    public void setName(String name) {
        this.name = name;
    }

    public String getLink() {
        return link;
    }

    public void setLink(String link) {
        this.link = link;
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
