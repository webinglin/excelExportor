package com.piedra.excel;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.util.*;

import com.piedra.excel.bean.ExcelRow;
import com.piedra.excel.bean.ExcelRowForMultiHeaders;
import org.junit.Test;

import com.piedra.excel.util.ExcelExportor;

/**
 * @Description: Excel导出工具  例子程序
 * @Creator：linwb 2014-12-19
 */
public class ExcelExportorTest {
    public static void main(String[] args) {
        new ExcelExportorTest().testSingleHeader();
        
        new ExcelExportorTest().testMulHeaders();
    }

    /**
     * 测试单表头，多个sheet表格
     */
    @Test
    public void testSingleHeaderMultiSheet(){
        OutputStream out = null;
        try {
            out = new FileOutputStream(new File("C://EXCEL-EXPORT-TEST-MULTI.xls"));

            List<Object> stus = new ArrayList<Object>();
            for(int i=0; i<11120; i++){
                stus.add(new ExcelRow());
            }

            Map<String,List<Object>> multiSheetDatas = new HashMap<String,List<Object>>();
            multiSheetDatas.put("表格一",stus);
            multiSheetDatas.put("第二个表格",stus);
            multiSheetDatas.put("Thrid Sheet",stus);

            new ExcelExportor<Object>().exportExcel(multiSheetDatas,out);

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
            Map<String,List<ExcelRow>> map = new HashMap<String,List<ExcelRow>>();
            map.put("测试单表头",stus);
            new ExcelExportor<ExcelRow>().exportExcel(map, out);
            
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

            Map<String,List<ExcelRowForMultiHeaders>> map = new HashMap<String,List<ExcelRowForMultiHeaders>>();
            map.put("测试多表头",stus);
            new ExcelExportor<ExcelRowForMultiHeaders>().exportExcel(map, out);

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


