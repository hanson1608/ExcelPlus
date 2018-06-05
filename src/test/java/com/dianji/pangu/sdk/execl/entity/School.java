/*
 * Copyright (C), 2018-2018, 深圳点积科技有限公司
 * FileName: School
 * Author:   hanson
 * Date:   2018/5/19 9:25
 * Since: 1.0.0
 */
package com.dianji.pangu.sdk.execl.entity;

import com.dianji.pangu.sdk.excel.common.Excel;

import javax.validation.constraints.NotNull;

/**
 * 学校
 *
 * @author hanson
 * @since 1.0.0
 * 2018/5/19
 */
@SuppressWarnings("unused")
public class School {
    @Excel(title="校长")
    private SchoolMaster schoolMaster;
    /**
     * 学校名称
     */
    @NotNull(message = "学校名称不能为空")
    @Excel(title = "学校名称", width = 30)
    private String schoolName;
    /**
     * 学校地址
     */
    @Excel(title = "学校地址", width = 30)
    private String address;

    public String getSchoolName() {
        return schoolName;
    }

    public void setSchoolName(String schoolName) {
        this.schoolName = schoolName;
    }

    public String getAddress() {
        return address;
    }

    public void setAddress(String address) {
        this.address = address;
    }

    public SchoolMaster getSchoolMaster() {
        return schoolMaster;
    }

    public void setSchoolMaster(SchoolMaster schoolMaster) {
        this.schoolMaster = schoolMaster;
    }

    @Override
    public String toString(){
        //对于字符串，null变成""
        StringBuilder sb=new StringBuilder();
        if(schoolName !=null){
            sb.append("学校名称:").append(schoolName);
        }else{
            sb.append("学校名称:"+"");
        }
        if(address!=null){
            sb.append("学校地址:").append(address);
        }else{
            sb.append("学校地址:"+"");
        }
        if(schoolMaster !=null){
            sb.append(schoolMaster.toString());
        }
        return sb.toString();
    }
}
