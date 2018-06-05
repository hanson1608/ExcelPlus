/*
 * Copyright (C), 2018-2018, 深圳点积科技有限公司
 * FileName: Student
 * Author:   hanson
 * Date:   2018/5/19 9:23
 * Since: 1.0.0
 */
package com.dianji.pangu.sdk.execl.entity;

import com.dianji.pangu.sdk.excel.common.Excel;

/**
 * 测试类
 *
 * @author hanson
 * @since 1.0.0
 * 2018/5/19
 */
@Excel(title="学生")
@SuppressWarnings("unused")
public class Student extends Person{
    @Excel(title="年级",format="##0")
    private int grade;
    @Excel(title="班级",format="##0")
    private long classNo;
    @Excel(title="成绩",format="##0.0")
    private double score;
    /**
     * 注明子类下面的字段进入excel列
     */
    @Excel(title="学校")
    private School school;

    public School getSchool() {
        return school;
    }

    public void setSchool(School school) {
        this.school = school;
    }

    public int getGrade() {
        return grade;
    }

    public void setGrade(int grade) {
        this.grade = grade;
    }

    public long getClassNo() {
        return classNo;
    }

    public void setClassNo(long classNo) {
        this.classNo = classNo;
    }

    public double getScore() {
        return score;
    }

    public void setScore(double score) {
        this.score = score;
    }

    @Override
    public String toString(){
        String personStr = super.toString();
        String schoolStr = "";
        if (school!=null){
            schoolStr = school.toString();
        }
        return personStr+schoolStr+"年级:"+grade+"班级:"+classNo+"成绩:"+score;
    }

}
