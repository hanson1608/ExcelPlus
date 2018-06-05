/*
 * Copyright (C), 2018-2018, 深圳点积科技有限公司
 * FileName: Excel
 * Author:   luoguanghan
 * Date:     2018/4/12 14:36
 *
 * @since 1.0.0
 */
package com.dianji.pangu.sdk.excel.common;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

/**
 *  Excel注解，用以生成Excel表格文件，标在类上对应sheet名，标在字段上对应列标题
 * @author luoguanghan
 * @since 1.0.0
 * Date 2018/4/12
 */
@Retention(RetentionPolicy.RUNTIME)
@Target({ElementType.FIELD,ElementType.TYPE})
public @interface Excel {
    /**
     * 列标题 缺省为""
     */
    String title() default "";
    /**
     * excel的列宽度 缺省为20
     */
    int width() default 20;
    /**
     * 该字段数据在excel里展示格式
     */
    String  format() default "";


}
