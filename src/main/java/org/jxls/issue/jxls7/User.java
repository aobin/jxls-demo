package org.jxls.issue.jxls7;

import org.jxls.issue.jxls7.Data;

import java.util.ArrayList;
import java.util.List;

public class User {

    private int account;
    private String name;

    private List<Data> datas = new ArrayList<Data>();

    public int getAccount() {
        return account;
    }

    public void setAccount(int account) {
        this.account = account;
    }

    public String getName() {
        return name;
    }

    public void setName(String name) {
        this.name = name;
    }

    public List<Data> getDatas() {
        return datas;
    }

    public void setDatas(List<Data> datas) {
        this.datas = datas;
    }

}
