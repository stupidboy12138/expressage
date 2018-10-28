package com.stylefeng.guns.modular.system.excel;

import java.io.Serializable;
import java.util.List;

/**
 * Created by zouLu on 2017-12-14.
 */

public class ExcelData implements Serializable {

    private static final long serialVersionUID = 4444017239100620999L;

    // 表头
    private List<String> titles;

    // 数据
    private List<List<Object>> rows;

    // 页签名称
    private String name;

    private int exrow;

    public List<String> getTitles() {
        return titles;
    }

    public void setTitles(List<String> titles) {
        this.titles = titles;
    }

    public List<List<Object>> getRows() {
        return rows;
    }

    public void setRows(List<List<Object>> rows) {
        this.rows = rows;
    }

    public String getName() {
        return name;
    }

    public void setName(String name) {
        this.name = name;
    }

    public int getExrow() {
        return exrow;
    }

    public void setExrow(int exrow) {
        this.exrow = exrow;
    }
}

