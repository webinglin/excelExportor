package example;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;

import org.junit.Test;

import com.piedra.excel.annotation.ExcelExport;
import com.piedra.excel.util.ExcelExportor;

/**
 * @Description: Excel导出工具  例子程序
 * @Creator：linwb 2014-12-19
 */
public class ExcelExportorExample {
    public static void main(String[] args) {
        new ExcelExportorExample().testSingleHeader();
        
        new ExcelExportorExample().testMulHeaders();
    }
    
    /**
     * @Description: 测试单表头
     * @History
     *     1. 2014-12-19 linwb 创建方法
     */
    @Test
    public void testSingleHeader(){
        OutputStream out = null;
        try {
            out = new FileOutputStream(new File("C://EXCEL-EXPORT-TEST.xls"));
            
            List<ExcelRow> stus = new ArrayList<ExcelRow>();
            for(int i=0; i<11120; i++){
                stus.add(new ExcelRow());
            }
            
            new ExcelExportor<ExcelRow>().exportExcel("测试单表头",  stus, out);
            
            System.out.println("excel导出成功！");
        } catch (Exception e) {
           e.printStackTrace();
        } finally {
            if(out!=null){
                try {
                    out.close();
                } catch (IOException e) {
                    //Ignore..
                } finally{
                    out = null;
                }
            }
        }
    }
    
    
    /**
     * @Description: 测试多表头
     * @History
     *     1. 2014-12-19 linwb 创建方法
     */
    @Test
    public void testMulHeaders(){
        OutputStream out = null;
        try {
            out = new FileOutputStream(new File("C://EXCEL-EXPORT-TEST-MULTIHEADER.xls"));
            
            List<ExcelRowForMultiHeaders> stus = new ArrayList<ExcelRowForMultiHeaders>();
            for(int i=0; i<1120; i++){
                stus.add(new ExcelRowForMultiHeaders());
            }
            new ExcelExportor<ExcelRowForMultiHeaders>().exportExcel("测试多表头",  stus, out);
            
            System.out.println("excel导出成功！");
        } catch (Exception e) {
           e.printStackTrace();
        } finally {
            if(out!=null){
                try {
                    out.close();
                } catch (IOException e) {
                    //Ignore..
                } finally{
                    out = null;
                }
            }
        }
    }
}


/**
 * @Description: Excel的一行对应的JavaBean类
 * @Creator：linwb 2014-12-19
 */
class ExcelRow {
    
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

/**
 * @Description: Excel的一行对应的JavaBean类
 * @Creator：linwb 2014-12-19
 */
class ExcelRowForMultiHeaders {
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
