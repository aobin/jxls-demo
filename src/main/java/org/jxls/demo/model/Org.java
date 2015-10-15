package org.jxls.demo.model;

import java.util.ArrayList;
import java.util.List;
import java.util.Random;

/**
 * Created by Leonid Vysochyn on 10-Oct-15.
 */
public class Org {
    String name;
    List<Department> departments;

    public Org(String name) {
        this.name = name;
    }

    public static List<Org> generate(int orgCount, int depCount){
        List<Org> orgs = new ArrayList<Org>();
        Random random = new Random(System.currentTimeMillis());
        for(int index = 0; index < orgCount; index++){
            Org org = new Org("Org " + index);
            org.setDepartments(Department.generate(depCount, 2 + random.nextInt(5)));
            orgs.add(org);
        }
        return orgs;
    }

    public String getName() {
        return name;
    }

    public void setName(String name) {
        this.name = name;
    }

    public List<Department> getDepartments() {
        return departments;
    }

    public void setDepartments(List<Department> departments) {
        this.departments = departments;
    }
}
