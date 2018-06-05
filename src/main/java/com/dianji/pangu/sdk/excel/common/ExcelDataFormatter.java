/*
 * Copyright (C), 2018-2018, 深圳点积科技有限公司
 * FileName: ExcelDataFormatter
 * Author:   luoguanghan
 * Date:     2018/4/12 14:46
 *
 * @since 1.0.0
 */
package com.dianji.pangu.sdk.excel.common;

import java.util.HashMap;
import java.util.Map;

/**
 * Excel数据导入导出格式化<br>
 * @author luoguanghan
 * @since 1.0.0
 * Date 2018/4/12
 */
public class ExcelDataFormatter {

    private Map<String,Map<String,String>> formatter=new HashMap<>();

    public void set(String key,Map<String,String> map){
        formatter.put(key, map);
    }

    public Map<String,String> get(String key){
        return formatter.get(key);
    }
}
