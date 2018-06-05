/*
 * Copyright (C), 2018-2018, 深圳点积科技有限公司
 * FileName: SheetConfig
 * Author:   hanson
 * Date:   2018/5/18 17:13
 * Since: 1.0.0
 */
package com.dianji.pangu.sdk.excel.common;

import java.util.Map;

/**
 * Excel 的Sheet的配置类
 *
 * @author hanson
 * @since 1.0.0  2018/5/18
 */
@SuppressWarnings("unused")
public class SheetConfig {
    /**
     * 每页Sheet 最大的行数
     */
    public static final int SHEET_MAX_ROWS = 65535;
    /**
     * 缺省整数格式
     */
    private static final String DEFAULT_INTEGER_FORMAT = "#,##0";
    /**
     * 缺省小数格式，两位小数
     */
    private static final String DEFAULT_DOUBLE_FORMAT = "#,##0.00";
    /**
     * 缺省日期格式
     */
    private static final String DEFAULT_DATE_FORMAT = "yyyy-MM-dd HH:mm:ss";
    /**
     * 缺省列宽单位256
     */
    private static final int DEFAULT_CELL_UNIT = 256;
    /**
     * 标题所在的行号
     */
    private int titleRow = 0;
    /**
     * 数据的开始行
     */
    private int dataBeginRow = 1;
    /**
     * 数据每个Cell的基本单元宽度，cell的实际宽度为根据bean里定义的宽度*这个的宽度
     */
    private int cellUnitWidth = DEFAULT_CELL_UNIT;
    /**
     * 整数类型的格式
     */
    private String integerFormat = DEFAULT_INTEGER_FORMAT;
    /**
     * 小数类型的格式
     */
    private String doubleFormat = DEFAULT_DOUBLE_FORMAT;
    /**
     * 日期时间类型的格式
     */
    private String dateFormat = DEFAULT_DATE_FORMAT;
    /**
     * 特殊数据映射转换，主要针对布尔和整型
     */
    private ExcelDataFormatter edf;
    /**
     * 支持非注解方式，key:字段名和value:excel列名对应，字段名是字段完整名，格式：fieldFullName = parentFieldFullName.fieldName。父字段必须要在map里
     */
    private Map<String,String> fieldNameMap =null;
    public int getTitleRow() {
        return titleRow;
    }

    public void setTitleRow(int titleRow) {
        if(titleRow>=0) {
            this.titleRow = titleRow;
        }
    }

    public int getDataBeginRow() {
        return dataBeginRow;
    }

    public void setDataBeginRow(int dataBeginRow) {
        if(dataBeginRow>0){
            this.dataBeginRow = dataBeginRow;
        }
    }

    public int getCellUnitWidth() {
        return cellUnitWidth;
    }

    public void setCellUnitWidth(int cellUnitWidth) {
        this.cellUnitWidth = cellUnitWidth;
    }

    public String getIntegerFormat() {
        return integerFormat;
    }

    public void setIntegerFormat(String integerFormat) {
        this.integerFormat = integerFormat;
    }

    public String getDoubleFormat() {
        return doubleFormat;
    }

    public void setDoubleFormat(String doubleFormat) {
        this.doubleFormat = doubleFormat;
    }

    public String getDateFormat() {
        return dateFormat;
    }

    public void setDateFormat(String dateFormat) {
        this.dateFormat = dateFormat;
    }

    public ExcelDataFormatter getEdf() {
        return edf;
    }

    public void setEdf(ExcelDataFormatter edf) {
        this.edf = edf;
    }

    public Map<String, String> getFieldNameMap() {
        return fieldNameMap;
    }

    public void setFieldNameMap(Map<String, String> fieldNameMap) {
        this.fieldNameMap = fieldNameMap;
    }
}
