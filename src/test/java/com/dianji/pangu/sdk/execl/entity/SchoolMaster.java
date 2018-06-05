/*
 * Copyright (C), 2018-2018, 深圳点积科技有限公司
 * FileName: Master
 * Author:   hanson
 * Date:   2018/5/22 18:25
 * Since: 1.0.0
 */
package com.dianji.pangu.sdk.execl.entity;

import com.dianji.pangu.sdk.excel.common.Excel;

/**
 * @author hanson
 * @since 1.0.0
 * 2018/5/22
 */
@SuppressWarnings("unused")
public class SchoolMaster {
    @Excel(title = "校长姓名")
    private String name;
    @Excel(title="校长性别")
    private String sex;

    public SchoolMaster(){

    }
    public SchoolMaster(String name, String sex) {
        this.name = name;
        this.sex = sex;
    }

    public String getName() {
        return name;
    }

    public void setName(String name) {
        this.name = name;
    }

    public String getSex() {
        return sex;
    }

    public void setSex(String sex) {
        this.sex = sex;
    }
    @Override
    public String toString(){
        //对于字符串，null变成""
        StringBuilder sb=new StringBuilder();
        if(name!=null){
            sb.append("校长姓名:").append(name);
        }else{
            sb.append("校长姓名:"+"");
        }
        if(sex!=null){
            sb.append("校长性别:").append(sex);
        }else{
            sb.append("校长性别:"+"未知");
        }
        return sb.toString();
    }

}
