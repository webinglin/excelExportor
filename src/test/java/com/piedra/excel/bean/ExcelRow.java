package com.piedra.excel.bean;

import com.piedra.excel.annotation.ExcelExport;

import java.util.Date;

/**
 * @Description: Excel的一行对应的JavaBean类
 * @Creator：linwb 2014-12-19
 */
public class ExcelRow {

    @ExcelExport(header="姓名",colWidth=50)
    private String name="AAAAAAAAAAASSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSSS";
    @ExcelExport(header="年龄")
    private int age=80;

    /** 这个属性没别注解,那么将不会出现在导出的excel文件中*/
    private String clazz="SSSSSSSSS";

    @ExcelExport(header="国家")
    private String country="RRRRRRR";
    @ExcelExport(header="城市")
    private String city="EEEEEEEEE";
    @ExcelExport(header="城镇")
    private String town="WWWWWWW";

    /** 这个属性没别注解,那么将不会出现在导出的excel文件中*/
    private String common="DDDDDDDD";

    /** 如果colWidth <= 0 那么取默认的 15 */
    @ExcelExport(header="出生日期",colWidth=-1)
    private Date birth = new Date();

    public ExcelRow(){
    }


    public String getName() {
        return name;
    }
    public void setName(String name) {
        this.name = name;
    }
    public int getAge() {
        return age;
    }
    public void setAge(int age) {
        this.age = age;
    }
    public String getClazz() {
        return clazz;
    }
    public void setClazz(String clazz) {
        this.clazz = clazz;
    }
    public String getCommon() {
        return common;
    }
    public void setCommon(String common) {
        this.common = common;
    }
    public Date getBirth() {
        return birth;
    }
    public void setBirth(Date birth) {
        this.birth = birth;
    }
    public String getCountry() {
        return country;
    }
    public void setCountry(String country) {
        this.country = country;
    }
    public String getCity() {
        return city;
    }
    public void setCity(String city) {
        this.city = city;
    }
    public String getTown() {
        return town;
    }
    public void setTown(String town) {
        this.town = town;
    }
}
