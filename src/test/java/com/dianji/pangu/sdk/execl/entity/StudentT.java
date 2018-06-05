/*
 * Copyright (C), 2018-2018, 深圳点积科技有限公司
 * FileName: StudentT
 * Author:   hanson
 * Date:   2018/5/27 15:27
 * Since: 1.0.0
 */
package com.dianji.pangu.sdk.execl.entity;

import com.dianji.pangu.sdk.excel.common.Excel;

/**
 * 测试泛型
 *
 * @author hanson
 * @since 1.0.0
 * 2018/5/27
 */
@SuppressWarnings("unused")
public class StudentT<T,V> {
    @Excel(title = "StudentT.x")
    private int x;
    @Excel(title = "StudentT.y")
    private Double y;
    @Excel(title = "StudentT.name")
    private String name;
    @Excel(title = "StudentT.T")
    private T obj;
    @Excel(title = "StudentT.V")
    private V vobj;

    public T getObj() {
        return obj;
    }

    public void setObj(T obj) {
        this.obj = obj;
    }

    public V getVobj() {
        return vobj;
    }

    public void setVobj(V vobj) {
        this.vobj = vobj;
    }

    public int getX() {
        return x;
    }

    public void setX(int x) {
        this.x = x;
    }

    public Double getY() {
        return y;
    }

    public void setY(Double y) {
        this.y = y;
    }

    public String getName() {
        return name;
    }

    public void setName(String name) {
        this.name = name;
    }
}
