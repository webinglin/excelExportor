package com.piedra.excel.bean;

import com.piedra.excel.annotation.ExcelExport;

import java.util.Date;

/**
 * @Description: Excel的一行对应的JavaBean类
 * @Creator：linwb 2014-12-19
 */
public class ExcelRowForMultiHeaders {
    @ExcelExport(header="姓名",colspan="1",rowspan="3")
    private String name="无名氏";

    @ExcelExport(header="省份,国家",colspan="1,5",rowspan="1,2")
    private String province="福建省";
    @ExcelExport(header="城市",colspan="1",rowspan="1")
    private String city="福建省";
    @ExcelExport(header="城镇",colspan="1",rowspan="1")
    private String town="不知何处";

    @ExcelExport(header="年龄,年龄和备注",colspan="1,2",rowspan="1,1")
    private int age=80;
    @ExcelExport(header="备注?",colspan="1",rowspan="1")
    private String common="我是备注,我是备注";

    @ExcelExport(header="我的生日",colspan="1",rowspan="3",datePattern="yyyy-MM-dd HH:mm:ss")
    private Date birth = new Date();

    /** 这个属性没别注解,那么将不会出现在导出的excel文件中*/
    private String clazz="我不会出现的,除非你给我 @ExcelExport 注解标记";

    public ExcelRowForMultiHeaders(){
    }

    public String getClazz() {
        return clazz;
    }
    public void setClazz(String clazz) {
        this.clazz = clazz;
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
    public String getProvince() {
        return province;
    }
    public void setProvince(String province) {
        this.province = province;
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
