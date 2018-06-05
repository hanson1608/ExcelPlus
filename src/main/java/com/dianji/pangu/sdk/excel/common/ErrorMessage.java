/*
 * Copyright (C), 2018-2018, 深圳点积科技有限公司
 * FileName: ErrorMessage
 * Author:   hanson
 * Date:   2018/5/20 10:55
 * Since: 1.0.0
 */
package com.dianji.pangu.sdk.excel.common;

/**
 * 错误信息
 *
 * @author hanson
 * @since 1.0.0
 * 2018/5/20
 */
@SuppressWarnings("unused")
public class ErrorMessage {
    public static final String NO_SHEET_DATA = "没有 Sheet 数据";
    public static final String NO_ENTITY_DATA = "没有实体数据";
    public static final String NO_TITLE_DATA = "没有标题数据";
    public static final String NO_CONSTRUCTOR_WITHOUT_PARA = "没有无参构造方法";
    public static final String TOO_MANY_ROWS = "太多行，超过了最大行数";
    public static final String VERIRY_ERROR = "校验错误";

    /**
     * 错误行
     */
    private int row;
    /**
     * 错误信息
     */
    private StringBuilder msg;

    public ErrorMessage(int row, StringBuilder msg) {
        this.row = row;
        this.msg = msg;
    }

    public int getRow() {
        return row;
    }

    public void setRow(int row) {
        this.row = row;
    }

    public StringBuilder getMsg() {
        return msg;
    }

    public void setMsg(StringBuilder msg) {
        this.msg = msg;
    }
}
