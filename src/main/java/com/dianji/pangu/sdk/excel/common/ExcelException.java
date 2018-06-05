/*
 * Copyright (C), 2018-2018, 深圳点积科技有限公司
 * FileName: ExcelException
 * Author:   hanson
 * Date:   2018/5/22 20:10
 * Since: 1.0.0
 */
package com.dianji.pangu.sdk.excel.common;

/**
 * Excel解析的异常
 *
 * @author hanson
 * @since 1.0.0
 * 2018/5/22
 */
public class ExcelException extends Exception {
    private static final long serialVersionUID = 1L;

    public ExcelException(String message){
        super(message);
    }
}
