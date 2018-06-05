/*
 * Copyright (C), 2018-2018, 深圳点积科技有限公司
 * FileName: ValidatorUtil
 * Author:   hanson
 * Date:   2018/5/19 11:40
 * Since: 1.0.0
 */
package com.dianji.pangu.sdk.excel.util;

import java.util.*;
import javax.validation.ConstraintViolation;
import javax.validation.Validation;
import javax.validation.Validator;
import javax.validation.groups.Default;

/**
 * 数据校验工具类
 *
 * @author hanson
 * @since 1.0.0
 * 2018/5/19
 */
@SuppressWarnings("unused")
public class ValidatorUtil {

    private ValidatorUtil(){
    }
    /**
     * 校验器
     */
    private static Validator validator = Validation.buildDefaultValidatorFactory().getValidator();

    public static <T> Map<String,StringBuilder> validate2(T obj){
        Map<String,StringBuilder> errorMap = null;
        Set<ConstraintViolation<T>> set = validator.validate(obj,Default.class);
        if(set != null && !set.isEmpty() ){
            errorMap = new HashMap<>(16);
            String property;
            for(ConstraintViolation<T> cv : set){
                //这里循环获取错误信息，可以自定义格式
                property = cv.getPropertyPath().toString();
                if(errorMap.get(property) != null){
                    errorMap.get(property).append(",").append(cv.getMessage());
                }else{
                    StringBuilder sb = new StringBuilder();
                    sb.append(cv.getMessage());
                    errorMap.put(property, sb);
                }
            }
        }
        return errorMap;
    }

    /**
     * 校验对象，返回各个字段错误校验结果信息
     * @param obj 实体类
     * @return  List<String> 校验错误的消息列表
     */
    public static  <T> List<String> validate(T obj) {
        Set<ConstraintViolation<T>> constraintViolations = validator.validate(obj);
        List<String> errorMessageList = new ArrayList<>();
        if(constraintViolations != null && !constraintViolations.isEmpty() ){
            for (ConstraintViolation<T> constraintViolation : constraintViolations) {
                errorMessageList.add(constraintViolation.getMessage());
            }
        }
        return errorMessageList;
    }

}
