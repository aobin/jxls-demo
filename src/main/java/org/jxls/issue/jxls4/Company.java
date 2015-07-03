package org.jxls.issue.jxls4;

import java.util.ArrayList;
import java.util.List;

/**
 *
 * @author pernik
 */
public class Company {
    private Long id;
    private String name;
    private List<Person> girls = new ArrayList<>();
    private List<Person> boys = new ArrayList<>();
    
    public Company(Long id, String name) {
        this.id = id;
        this.name = name;
    }
    
    public Company addBoy(String p) {
        this.boys.add(new Person(p));
        return this;
    }
    public Company addGirl(String p) {
        this.girls.add(new Person(p));
        return this;
    }

    /**
     * @return the id
     */
    public Long getId() {
        return id;
    }

    /**
     * @param id the id to set
     */
    public void setId(Long id) {
        this.id = id;
    }

    /**
     * @return the name
     */
    public String getName() {
        return name;
    }

    /**
     * @param name the name to set
     */
    public void setName(String name) {
        this.name = name;
    }

    /**
     * @return the girls
     */
    public List<Person> getGirls() {
        return girls;
    }

    /**
     * @param girls the girls to set
     */
    public void setGirls(List<Person> girls) {
        this.girls = girls;
    }

    /**
     * @return the boys
     */
    public List<Person> getBoys() {
        return boys;
    }

    /**
     * @param boys the boys to set
     */
    public void setBoys(List<Person> boys) {
        this.boys = boys;
    }
}
