package com.piedra.excel.bean;

/**
 * @Description: Excel导出表格的头部信息
 * @Creator：linwb 2014-12-20
 */
public class ExcelHeader {
    /** 表头的文字描述*/
    private String[] headerArr;
    /** 多表头的时候配置的各个单元格的列    详见文档说明*/
    private String[] colspanArr;
    /** 多表头的时候配置的各个单元格的行   详见文档说明*/
    private String[] rowspanArr;
    /** 列宽*/
    private int columnWidth;
    /** 如果是日期类型的那么该值将被采用*/
    private String datePattern;
    
    public ExcelHeader(){
    }
    
    public ExcelHeader(String[] headerArr, String[] rowspanArr, String[] colspanArr, int columnWidth, String datePattern){
        this.headerArr = headerArr;
        this.rowspanArr = rowspanArr;
        this.colspanArr = colspanArr;
        this.columnWidth = columnWidth;
        this.datePattern = datePattern;
    }
    
    public String[] getHeaderArr() {
        return headerArr;
    }
    public void setHeaderArr(String[] headerArr) {
        this.headerArr = headerArr;
    }
    public String[] getColspanArr() {
        return colspanArr;
    }
    public void setColspanArr(String[] colspanArr) {
        this.colspanArr = colspanArr;
    }
    public String[] getRowspanArr() {
        return rowspanArr;
    }
    public void setRowspanArr(String[] rowspanArr) {
        this.rowspanArr = rowspanArr;
    }
    public int getColumnWidth() {
        return columnWidth;
    }
    public void setColumnWidth(int columnWidth) {
        this.columnWidth = columnWidth;
    }
    public String getDatePattern() {
        return datePattern;
    }
    public void setDatePattern(String datePattern) {
        this.datePattern = datePattern;
    }
}
