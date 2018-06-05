/*
 * Copyright (C), 2018-2018, 深圳点积科技有限公司
 * FileName: FieldTree
 * Author:   hanson
 * Date:   2018/5/20 11:54
 * Since: 1.0.0
 */
package com.dianji.pangu.sdk.excel.common;

import java.lang.reflect.Field;
import java.util.ArrayList;
import java.util.List;

/**
 * 字段和列对应关系树
 * 保存字段和列号对应关系的树。考虑到支持非基本类型和包装类的excel转换，用map保存字段和列的对应关系不够，必须采用树进行保存。
 * @author hanson
 * @since 1.0.0
 * 2018/5/20
 */
public class FieldTree {
    /**
     * 标志实体类中对应的Field
     */
    private Field field;
    /**
     * 记录fiedl对应的类，一般情况下和getType()一样，泛型的时候这里填写的是实际的类
     */
    private Class<?> fieldType;
    /**
     * 字段的完整名字，fieldName.fieldName，支持字段非基本类型
     */
    private String fieldFullName="";
    /**
     * 是否是基本类（java基本类型和包装类型），基本类没有儿子树
     */
    private boolean baseClassType;
    /**
     * 对应的列的标题名
     */
    private String colTitleName;
    /**
     * 列宽度
     */
    private int  colWidth;
    /**
     * 该列的格式
     */
    private String colFormat;
    /**
     * 对应当前excel sheet里的列的index
     */
    private int colIndex;
    /**
     * 当前字段为非基本类型或包装类时，该字段类下属的字段树
     */
    private List<FieldTree> children = new ArrayList<>();

    public Field getField() {
        return field;
    }

    public void setField(Field field) {
        this.field = field;
    }

    public boolean isBaseClassType() {
        return baseClassType;
    }

    public void setBaseClassType(boolean baseClassType) {
        this.baseClassType = baseClassType;
    }

    public int getColIndex() {
        return colIndex;
    }

    public void setColIndex(int colIndex) {
        this.colIndex = colIndex;
    }

    public String getColTitleName() {
        return colTitleName;
    }

    public void setColTitleName(String colTitleName) {
        this.colTitleName = colTitleName;
    }

    public List<FieldTree> getChildren() {
        return children;
    }

    public void addChild(FieldTree child) {
        children.add(child);
    }

    public int getColWidth() {
        return colWidth;
    }

    public void setColWidth(int colWidth) {
        this.colWidth = colWidth;
    }

    public String getColFormat() {
        return colFormat;
    }

    public void setColFormat(String colFormat) {
        this.colFormat = colFormat;
    }

    public String getFieldFullName() {
        return fieldFullName;
    }

    public void setFieldFullName(String fieldFullName) {
        this.fieldFullName = fieldFullName;
    }

    public Class<?> getFieldType() {
        return fieldType;
    }

    public void setFieldType(Class<?> fieldType) {
        this.fieldType = fieldType;
    }
}
