/*
 * Copyright (C), 2018-2018, 深圳点积科技有限公司
 * FileName: FieldConfig
 * Author:   hanson
 * Date:   2018/5/24 8:55
 * Since: 1.0.0
 */
package com.dianji.pangu.sdk.excel.common;

/**
 * 字段配置，暂未使用，读的时候暂时简单化不需要，写的时候需要
 *
 * @author hanson
 * @since 1.0.0
 * 2018/5/24
 */
@SuppressWarnings("unused")
public class FieldConfig {
    /**
     * 字段完整名称
     */
    private String fieldFullName;
    //对应注解里的参数
    /**
     * 对应的excel列名
     */
    private String colName="";
    /**
     * 宽度
     */
    private  int  width=20;
    /**
     * 格式
     */
    private String format="";

    public String getFieldFullName() {
        return fieldFullName;
    }

    public void setFieldFullName(String fieldFullName) {
        this.fieldFullName = fieldFullName;
    }

    public String getColName() {
        return colName;
    }

    public void setColName(String colName) {
        this.colName = colName;
    }

    public int getWidth() {
        return width;
    }

    public void setWidth(int width) {
        this.width = width;
    }

    public String getFormat() {
        return format;
    }

    public void setFormat(String format) {
        this.format = format;
    }
}
