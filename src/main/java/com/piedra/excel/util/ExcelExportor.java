package com.piedra.excel.util;

import java.io.IOException;
import java.io.OutputStream;
import java.lang.annotation.Annotation;
import java.lang.reflect.Field;
import java.lang.reflect.Method;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.HashMap;
import java.util.Iterator;
import java.util.List;
import java.util.Map;

import org.apache.commons.lang3.StringUtils;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.hssf.usermodel.HSSFRichTextString;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.util.CellRangeAddress;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import com.piedra.excel.Constants;
import com.piedra.excel.annotation.ExcelExport;
import com.piedra.excel.bean.ExcelHeader;

/**
 * @Description: Excel导出的通用类
 * @Creator：linwb 2014-12-19
 */
public class ExcelExportor<T> {
    
    private static final Logger logger = LoggerFactory.getLogger(ExcelExportor.class);
    
    /** 导出excel过程中,出错的数据行*/
    private List<T> errorDatas = new ArrayList<T>();
    /** 表头的配置信息*/
    private Map<String,ExcelHeader> headerMap = new HashMap<String,ExcelHeader>();
    
/* ****************************************************************************
 *          个性化定制要导出的数据样式
 * ****************************************************************************/
    
    /** 每一个sheet表格默认可以导出多少行*/
    private int maxSheetRows = 5000;
    
    /** 导出的列的头部 左边框*/
    private Short headerBorderLeft = HSSFCellStyle.BORDER_THIN;
    /** 导出的列的头部 右边框*/
    private Short headerBorderRight = HSSFCellStyle.BORDER_THIN;
    /** 导出的列的头部 底部边框*/
    private Short headerBorderBottom = HSSFCellStyle.BORDER_THIN;
    /** 导出的列的头部 顶部边框*/
    private Short headerBorderTop = HSSFCellStyle.BORDER_THIN;
    /** 导出的列的头部 文字对齐方式*/
    private Short headerCellTextAlign = HSSFCellStyle.ALIGN_CENTER;
    /** 导出的列的头部 文字垂直对齐方式*/
    private Short headerCellVehicleAlign = HSSFCellStyle.VERTICAL_CENTER;
    /** 导出的列的头部 背景颜色*/
    private Short headerCellBackgroundColor = HSSFColor.WHITE.index;
    
    /** 表头字体颜色*/
    private Short headerFontColor = HSSFColor.VIOLET.index;
    /** 表头字体的高度，即大小*/
    private Short headerFontHeight = 12;
    /** 表头字体的粗细*/
    private Short headerFontWeight = HSSFFont.BOLDWEIGHT_BOLD;
    
    /** 单元格的字体颜色*/
    private Short cellFontColor = HSSFColor.BLUE.index;
    /** 导出的内容的左边框*/
    private Short contBorderLeft = HSSFCellStyle.BORDER_THIN;
    /** 导出的内容的右边框*/
    private Short contBorderRight = HSSFCellStyle.BORDER_THIN;
    /** 导出的内容的底部边框*/
    private Short contBorderBottom = HSSFCellStyle.BORDER_THIN;
    /** 导出的内容的顶部边框*/
    private Short contBorderTop = HSSFCellStyle.BORDER_THIN;
    /** 导出的内容的文本水平对齐方式*/
    private Short contCellTextAlign = HSSFCellStyle.ALIGN_CENTER;
    /** 导出的内容的文本垂直对齐方式*/
    private Short contCellVehicleAlign = HSSFCellStyle.VERTICAL_CENTER;
    /** 导出的内容的背景颜色*/
    private Short contCellBackgroundColor = HSSFColor.WHITE.index;
    
    
    public ExcelExportor(){
    }
    
    
    /**
     * 通用的excel导出方法<br>
     * 注：<br>
     * 1. 当前版本不支持 boolean类型,byte[]类型  这些类型的请自己转成String或者int类型然后用 @ExcelExport 标记为要导出的字段<br>
     * 2. 每次要导出多少行,由开发者自行设定,此处不做限定; 但是每5000条数据的时候,将创建多个sheet表格<br>
     * 3. 如果要获取导出失败的数据,通过getErrorDatas()来获取<br>
     * 
     * @param headers   表格属性列名数组
     * @param dataset   需要显示的数据集合,集合里面的元素必须是JavaBean, 然后对于要导出的列 用 @ExcelExport annotation标记; 具体注解的使用,参照注解类的说明
     * @param out       与输出设备关联的流对象，可以将EXCEL文档导出到本地文件或者网络中
     *
     * @return Map&lt;Key, Msg&gt; <br>
     *             Key--表示成功与否的类型      Constants.SUCCESS_CODE 成功;     Constants.ERROR_CODE 错误;<br>
     *             Msg--表示成功或者失败的原因，为什么从该方法返回<br>
     * @History
     *     1. 2014-12-19 linwb 创建方法
     */
    public Map<String,String> exportExcel(String title, List<T> dataset, OutputStream out) {
        Map<String,String> map = new HashMap<String,String>();
        if(out==null){
            map.put(Constants.ERROR_CODE, "输出流为空,无法继续执行操作.");
            return map;
        }
        if(dataset==null || dataset.size()==0){
            map.put(Constants.ERROR_CODE, "没有任何要导出的数据,无法执行导出操作.");
            return map;
        }
        if(StringUtils.isBlank(title)){
            title = "通用excel导出工具";
        }
        
        //每一次要开始导出操作,清空errorDatas
        errorDatas.clear();
        
        //当前要导出的数据对象类型
        T beanType = (T) dataset.get(0);
        
        //1. 声明一个工作薄
        HSSFWorkbook workbook = new HSSFWorkbook();
        
        //2. 生成一个表格
        int sheets = 1;
        int dataSize = dataset.size();
        if(dataSize>maxSheetRows){
            sheets = (dataSize%maxSheetRows==0)?(dataSize/maxSheetRows):(dataSize/maxSheetRows+1);
        }
        
        //如果数据量太多生成多个表格
        int minIndex = 0,maxIndex=0;
        for(int i=0; i<sheets; i++){
            HSSFSheet sheet = workbook.createSheet(title+(i==0?"":""+i));
            
            //3. 生成表头
            int maxRows = generateHeader(workbook, sheet, beanType);
            //4. 将数据集的每一个数据项写到excel的每一行
            minIndex = i*maxSheetRows;
            if(dataSize/((i+1)*maxSheetRows) > 0){
                maxIndex = (i+1)*maxSheetRows;
            } else {
                maxIndex = dataSize;
            }
            
            List<T> dataRows = dataset.subList(minIndex, maxIndex);
            exportRowDatas(dataRows, workbook, sheet, maxRows);
        }
        
        //5. 将表格输出到指定的输出流
        try {
            workbook.write(out);
        } catch (IOException e) {
            String errorMsg = "生成excel表格失败";
            logger.error(errorMsg, e);
            map.put(Constants.ERROR_CODE, errorMsg);
            return map;
        }
       
        map.put(Constants.SUCCESS_CODE, "导出成功.");
        return map;
    }

    /**
     * @Description: 将数据集的每一个数据项写到excel的每一行
     * @param dataset
     * @param workbook
     * @param sheet
     *
     * @History
     *     1. 2014-12-19 linwb 创建方法
     */
    @SuppressWarnings({"unchecked" })
    private void exportRowDatas(List<T> dataset, HSSFWorkbook workbook, HSSFSheet sheet, int maxRows) {
        //生成数据单元格的字体
        HSSFFont contFont = workbook.createFont();
        contFont.setColor(cellFontColor);
        //生成文本单元格内容的样式
        HSSFCellStyle cellStyle = generateContStyle(workbook);
        
        HSSFRow row;
        //遍历集合数据，产生数据行
        Iterator<T> it = dataset.iterator();
        int index = maxRows;
        boolean isErrorRow = false;
        while (it.hasNext()) {
            row = sheet.createRow(index++);
            T rowData = (T) it.next();
            
            isErrorRow = false;
            //根据JavaBean属性的先后顺序,动态调用getXxx()方法得到属性值
            Field[] fields = rowData.getClass().getDeclaredFields();
            for (int i = 0,colIndex=i, length=fields.length; i < length; i++) {
                Field field = fields[i];
                if(!isExcelExport(field.getAnnotations())){
                    continue ;
                }
                
                HSSFCell cell = row.createCell(colIndex++);
                cell.setCellStyle(cellStyle);
                
                String fieldName = field.getName();
                String getMethodName = "get" + fieldName.substring(0, 1).toUpperCase() + fieldName.substring(1);
                try {
                    Class tCls = rowData.getClass();
                    Method getMethod = tCls.getMethod(getMethodName, new Class[] {});
                    getMethod.setAccessible(true);
                    Object value = getMethod.invoke(rowData, new Object[] {});
                    if(value == null){
                        value = "";
                    }
                    //类型转换处理
                    String textValue = null;
                    if (value instanceof Date) {
                        Date date = (Date) value;
                        
                        ExcelHeader excelHeader = headerMap.get(""+(colIndex-1));//因为colIndex第几列已经增加过1了，所以此处减1获取正确的列
                        SimpleDateFormat sdf = new SimpleDateFormat(excelHeader.getDatePattern());
                        textValue = sdf.format(date);
                    } else {
                        //其它数据类型都当作字符串简单处理
                        textValue = value.toString();
                        
                        //指定部分字体的颜色?
                        
                    }
                    HSSFRichTextString richText = new HSSFRichTextString(textValue);
                    richText.applyFont(contFont);
                    cell.setCellValue(richText);
                    
                } catch (Exception e) {
                    e.printStackTrace();
                    logger.error("写入excel表格时出错", e);
                    isErrorRow = true;
                    break;
                } finally {
                    //清理资源
                }
            }
            if(isErrorRow){
                errorDatas.add(rowData);
            }
        }
    }
    
    /**
     * @Description: 生成表头数据  并返回表头共有多少行
     * @param workbook
     * @param sheet
     * @param beanType
     *
     * @History
     *     1. 2014-12-20 linwb 创建方法
     */
    private int generateHeader(HSSFWorkbook workbook, HSSFSheet sheet, T beanType) {
        //生成标题栏的样式
        HSSFCellStyle headerStyle = generateHeaderStyle(workbook);
        HSSFRow[] headerRows = null;
        int maxRow = 0;
        int headerColIndex = 0;
        Field[] fields = beanType.getClass().getDeclaredFields();
        for(Field field : fields ){
            Annotation[] annotions = field.getAnnotations();
            ExcelExport excelAnnotation = null;
            for(Annotation annotation : annotions){
                if(annotation instanceof ExcelExport){
                    excelAnnotation = (ExcelExport)annotation;
                    break;
                }
            }
            if(excelAnnotation == null){
                continue ;
            }
            
            String headerStr = excelAnnotation.header();
            String rowspanStr = excelAnnotation.rowspan();
            String colspanStr = excelAnnotation.colspan();
            int columnWidth = excelAnnotation.colWidth();
            String datePattern = excelAnnotation.datePattern();
            
            String[] headerArr = headerStr.split(",");
            String[] rowspanArr = rowspanStr.split(",");
            String[] colspanArr = colspanStr.split(",");
            
            try {
               if(headerArr.length!=rowspanArr.length || headerArr.length!=colspanArr.length){
                   throw new Exception("表头,列数,行数 配置不一致, 检查是否用英文的逗号分隔好?");
               }
               int curMaxRow = 0;
               int curRowspan = 1, curColspan = 1;
               for(int i=0,len=rowspanArr.length; i<len; i++){
                   //验证数值的合法性
                   curRowspan = Integer.parseInt(rowspanArr[i]);
                   curColspan = Integer.parseInt(colspanArr[i]);
                   if(curRowspan<=0 || curColspan<=0){
                       throw new Exception("占据的列数或者行数不能小于等于0. 请检查(colspan and rowspan should be greater than 0).");
                   }
                   if(i==0 && curColspan>1){
                       throw new Exception("最接近数据的列永远都只能是1(即你用逗号分割的那串colspan,第一个值必须是1; 如: '1,2,3'), 请确认您的表格设计是否合理");
                   }
                   curMaxRow += curRowspan;
               }
               //验证表头的行数,列数是否一致
               if(maxRow<curMaxRow){
                   maxRow = curMaxRow;
               }
            } catch(Exception e){
                throw new RuntimeException(e.getMessage() + "\n配置出错,请原谅我这么简单的提示.(检查建议: colspan,rowspan,header 检查用逗号分割后是否长度一致? col,row是否是数值?)");
            }
            //如果列宽小于等于0 那么设置成 15的列宽
            columnWidth = (columnWidth<=0 ? 15 : columnWidth);
            headerMap.put(""+headerColIndex++, new ExcelHeader(headerArr,rowspanArr,colspanArr,columnWidth,datePattern));
        }
        
        headerRows = new HSSFRow[maxRow];
        for(int i=0; i<maxRow; i++){
            headerRows[i] = sheet.createRow(i);
        }
        if(maxRow==0){
            throw new RuntimeException("表头行创建失败,请检查一下annotation的配置.");
        }
        
        for(int headerIndex=0,len=headerMap.size(); headerIndex<len; headerIndex++){
            ExcelHeader header = headerMap.get(""+headerIndex);
            String[] headerArr = header.getHeaderArr();
            String[] rowspanArr = header.getRowspanArr();
            String[] colspanArr = header.getColspanArr();
            
            //为表格的每一列设置列宽
            sheet.setColumnWidth(headerIndex, header.getColumnWidth()*256);
            
            //绘制表头
            int currentRow = maxRow-1 ;
            int minRowIndex = 0;
            for(int reverseRowIndex = headerArr.length-1,rowIndex=0; reverseRowIndex>=0; reverseRowIndex--,rowIndex++){
                if(rowIndex>0){
                    currentRow -= Integer.parseInt(rowspanArr[rowIndex-1]);//扣掉上一次的行跨度
                }
                
                if(Integer.parseInt(rowspanArr[rowIndex])>1 || Integer.parseInt(colspanArr[rowIndex])>1){
                    minRowIndex = currentRow-(Integer.parseInt(rowspanArr[rowIndex])-1);
                    sheet.addMergedRegion(new CellRangeAddress(minRowIndex,currentRow,
                            headerIndex,headerIndex+(Integer.parseInt(colspanArr[rowIndex])-1)));
                } else {
                    minRowIndex = currentRow;
                }
                
                //直接创建一个单元格.
                HSSFCell cell = headerRows[minRowIndex].createCell(headerIndex);
                cell.setCellStyle(headerStyle);
                
                //TODO headerArr[rowIndex] 变更为存放KEY 然后根据KEY 到数据库查询相应地区的表头的名称
                
                HSSFRichTextString richText = new HSSFRichTextString(headerArr[rowIndex]);
                cell.setCellValue(richText);
            }
        }
        
        return maxRow;
    }

    /**
     * Description: 生成表格的 头部的单元格样式
     * @param workbook
     * @return
     *
     * @History
     *     1. 2014-12-19 linwb 创建方法
     */
    private HSSFCellStyle generateHeaderStyle(HSSFWorkbook workbook) {
        //生成表格头部标题栏样式
        HSSFCellStyle headerStyle = workbook.createCellStyle();
        // 设置这些样式
        headerStyle.setFillForegroundColor(headerCellBackgroundColor);
        headerStyle.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
        
        headerStyle.setBorderBottom(headerBorderBottom);
        headerStyle.setBorderLeft(headerBorderLeft);
        headerStyle.setBorderRight(headerBorderRight);
        headerStyle.setBorderTop(headerBorderTop);
        
        headerStyle.setAlignment(headerCellTextAlign);
        headerStyle.setVerticalAlignment(headerCellVehicleAlign);
        
        // 生成字体
        HSSFFont font = workbook.createFont();
        font.setColor(headerFontColor);
        font.setFontHeightInPoints(headerFontHeight);
        font.setBoldweight(headerFontWeight);
        
        // 把字体应用到当前的样式
        headerStyle.setFont(font);
        
        return headerStyle;
    }

    /**
     * @Description: 生成excel表格 单元格内容的样式
     * @param workbook
     * @return
     *
     * @History
     *     1. 2014-12-19 linwb 创建方法
     */
    private HSSFCellStyle generateContStyle(HSSFWorkbook workbook) {
        HSSFCellStyle cellStyle = workbook.createCellStyle();
        cellStyle.setFillForegroundColor(contCellBackgroundColor);
        cellStyle.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
        
        cellStyle.setBorderBottom(contBorderBottom);
        cellStyle.setBorderLeft(contBorderLeft);
        cellStyle.setBorderRight(contBorderRight);
        cellStyle.setBorderTop(contBorderTop);
        cellStyle.setAlignment(contCellTextAlign);
        
        cellStyle.setVerticalAlignment(contCellVehicleAlign);
        // 生成字体
        HSSFFont font = workbook.createFont();
        font.setBoldweight(HSSFFont.BOLDWEIGHT_NORMAL);
        // 把字体应用到当前的样式
        cellStyle.setFont(font);
        
        return cellStyle;
    }

    
    /**
     * @Description: 判断是否包含 ExcelExport的annotation注解
     * @param annotations
     * @return  返回true表示要导出的列, false表示
     *
     * @History
     *     1. 2014-12-19 linwb 创建方法
     */
    private boolean isExcelExport(Annotation[] annotations){
        for(Annotation annotation : annotations){
            if(annotation instanceof ExcelExport){
                return true;
            }
        }
        return false;
    }
    
    
/* ****************************************************************************
 *          getters and setters...
 * ****************************************************************************/
   
    public Short getHeaderBorderLeft() {
        return headerBorderLeft;
    }
    
    public void setHeaderBorderLeft(Short headerBorderLeft) {
        this.headerBorderLeft = headerBorderLeft;
    }

    public Short getHeaderBorderRight() {
        return headerBorderRight;
    }

    public void setHeaderBorderRight(Short headerBorderRight) {
        this.headerBorderRight = headerBorderRight;
    }

    public Short getHeaderBorderBottom() {
        return headerBorderBottom;
    }

    public void setHeaderBorderBottom(Short headerBorderBottom) {
        this.headerBorderBottom = headerBorderBottom;
    }

    public Short getHeaderBorderTop() {
        return headerBorderTop;
    }

    public void setHeaderBorderTop(Short headerBorderTop) {
        this.headerBorderTop = headerBorderTop;
    }

    public Short getHeaderCellTextAlign() {
        return headerCellTextAlign;
    }

    public void setHeaderCellTextAlign(Short headerCellTextAlign) {
        this.headerCellTextAlign = headerCellTextAlign;
    }

    public Short getHeaderCellBackgroundColor() {
        return headerCellBackgroundColor;
    }

    public void setHeaderCellBackgroundColor(Short headerCellBackgroundColor) {
        this.headerCellBackgroundColor = headerCellBackgroundColor;
    }

    public Short getContBorderLeft() {
        return contBorderLeft;
    }

    public void setContBorderLeft(Short contBorderLeft) {
        this.contBorderLeft = contBorderLeft;
    }

    public Short getContBorderRight() {
        return contBorderRight;
    }

    public void setContBorderRight(Short contBorderRight) {
        this.contBorderRight = contBorderRight;
    }

    public Short getContBorderBottom() {
        return contBorderBottom;
    }

    public void setContBorderBottom(Short contBorderBottom) {
        this.contBorderBottom = contBorderBottom;
    }

    public Short getContBorderTop() {
        return contBorderTop;
    }

    public void setContBorderTop(Short contBorderTop) {
        this.contBorderTop = contBorderTop;
    }

    public Short getContCellTextAlign() {
        return contCellTextAlign;
    }

    public void setContCellTextAlign(Short contCellTextAlign) {
        this.contCellTextAlign = contCellTextAlign;
    }

    public Short getContCellBackgroundColor() {
        return contCellBackgroundColor;
    }

    public void setContCellBackgroundColor(Short contCellBackgroundColor) {
        this.contCellBackgroundColor = contCellBackgroundColor;
    }

    public List<T> getErrorDatas() {
        return errorDatas;
    }
    
    public void setErrorDatas(List<T> errorDatas) {
        this.errorDatas = errorDatas;
    }

    public Short getHeaderCellVehicleAlign() {
        return headerCellVehicleAlign;
    }

    public void setHeaderCellVehicleAlign(Short headerCellVehicleAlign) {
        this.headerCellVehicleAlign = headerCellVehicleAlign;
    }

    public Short getContCellVehicleAlign() {
        return contCellVehicleAlign;
    }

    public void setContCellVehicleAlign(Short contCellVehicleAlign) {
        this.contCellVehicleAlign = contCellVehicleAlign;
    }

    public int getMaxSheetRows() {
        return maxSheetRows;
    }
    public Short getHeaderFontColor() {
        return headerFontColor;
    }
    public void setHeaderFontColor(Short headerFontColor) {
        this.headerFontColor = headerFontColor;
    }
    public Short getHeaderFontHeight() {
        return headerFontHeight;
    }
    public void setHeaderFontHeight(Short headerFontHeight) {
        this.headerFontHeight = headerFontHeight;
    }
    public Short getHeaderFontWeight() {
        return headerFontWeight;
    }
    public void setHeaderFontWeight(Short headerFontWeight) {
        this.headerFontWeight = headerFontWeight;
    }
    public Short getCellFontColor() {
        return cellFontColor;
    }
    public void setCellFontColor(Short cellFontColor) {
        this.cellFontColor = cellFontColor;
    }
    
    public void setMaxSheetRows(int maxSheetRows) {
        if(maxSheetRows<=0){
            throw new RuntimeException("最大的表格行数不能小于0");
        }
        
        this.maxSheetRows = maxSheetRows;
    }
}
