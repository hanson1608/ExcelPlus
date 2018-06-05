/*
 * Copyright (C), 2018-2018, 深圳点积科技有限公司
 * FileName: Person
 * Author:   luoguanghan
 * Date:     2018/4/12 14:43
 *
 * @since 1.0.0
 */
package com.dianji.pangu.sdk.execl.entity;


import com.dianji.pangu.sdk.excel.common.Excel;
import java.math.BigDecimal;
import java.util.Date;

/**
 *
 * @author luoguanghan
 *  2018/4/12
 * @since 1.0.0
 */
@Excel(title="人员")
@SuppressWarnings("unused")
public class Person {
    @Excel(title = "姓名String", width = 30)
    private String name;


    @Excel(title = "年龄Long", width = 30)
    private Long age;

    @Excel(title ="密码String")
    private String password;

    @Excel(title = "Double测试" )
    private Double xx;

    @Excel(title = "日期yyyyMMdd",  width = 30, format = "yyyyMMdd")
    private Date yy;


    @Excel(title = "时间yyyyMMddHH:mm:ss",  width = 30, format = "yyyyMMddHH:mm:ss")
    private Date time;

    @Excel(title = "锁定Boolean")
    private Boolean locked;

   @Excel(title = "金额BigDecimal" )
    private BigDecimal db;

   @Excel(title = "FloatTest000",format = "0.000")
    private Float testFloat;

   @Override
   public String toString(){
       String nameStr = (name==null?"":name);
       String passwordStr=(password!=null?password:"");

       return "name: "+nameStr+" , " +"age: "+ age+" , "+"password: "+passwordStr+" , "+"xx: "+xx+"  ,"+"yy: "+yy+" , "+"time:"+time+", "+
               "locked: "+ locked+" , "+"db: "+db+" , "+"testFloat: "+testFloat+" ; ";
   }

    public Long getAge() {
        return age;
    }

    public void setAge(Long age) {
        this.age = age;
    }

    public Float getTestFloat() {
        return testFloat;
    }

    public void setTestFloat(Float testFloat) {
        this.testFloat = testFloat;
    }

    public String getName() {
        return name;
    }

    public void setName(String name) {
        this.name = name;
    }


    public String getPassword() {
        return password;
    }

    public void setPassword(String password) {
        this.password = password;
    }

    public Double getXx() {
        return xx;
    }

    public void setXx(Double xx) {
        this.xx = xx;
    }

    public Date getYy() {
        return yy;
    }

    public void setYy(Date yy) {
        this.yy = yy;
    }

    public Boolean getLocked() {
        return locked;
    }

    public void setLocked(Boolean locked) {
        this.locked = locked;
    }

    public BigDecimal getDb() {
        return db;
    }

    public void setDb(BigDecimal db) {
        this.db = db;
    }

    public Date getTime() {
        return time;
    }

    public void setTime(Date time) {
        this.time = time;
    }


}
