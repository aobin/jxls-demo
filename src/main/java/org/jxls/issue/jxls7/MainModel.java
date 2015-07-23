package org.jxls.issue.jxls7;

import java.util.ArrayList;
import java.util.List;

public class MainModel {

    private List<String> months = new ArrayList<String>();
    private List<User> users = new ArrayList<User>();

    public List<String> getMonths() {
        return months;
    }

    public void setMonths(List<String> months) {
        this.months = months;
    }

    public List<User> getUsers() {
        return users;
    }

    public void setUsers(List<User> users) {
        this.users = users;
    }

}
