package com.piedra.excel.annotation;

import java.lang.annotation.Documented;
import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

/**
 * @Description: excel导出的annotation标记类, 在字段上标记了该annotation之后表示该字段是要导出的
 * @Creator：linwb 2014-12-19
 */
@Retention(RetentionPolicy.RUNTIME)
@Target(ElementType.FIELD)
@Documented
public @interface ExcelExport {
    
    /**
     * @Description: 表示 导出的excel的标题, 多值用逗号隔开
     * @return
     * @History
     *     1. 2014-12-19 linwb 创建方法
     */
    String header() default "显以列名";
    
    /**
     * @Description: 导出的excel跨多少行  逗号隔开  如果是多表头的形式,都从最底下的算起  最靠近数据的那一行表头作为标记的地方
     * @return
     *
     * @History
     *     1. 2014-12-19 linwb 创建方法
     */
    String rowspan() default "1";
    
    /**
     * @Description: 导出的excel 跨多少列  逗号隔开    如果是多表头的形式,从最底下的算起  最靠近数据的那一行表头作为标记的地方
     * 
     * 列的第一个值拥有都是1  逗号分割之后,逗号后面的值可以是大于1的.但是第一个逗号之前的值必须是1
     * 
     * @return
     *
     * @History
     *     1. 2014-12-19 linwb 创建方法
     */
    String colspan() default "1";
    
    /**
     * @Description: 每一列的列宽
     * @return
     *
     * @History
     *     1. 2014-12-19 linwb 创建方法
     */
    int colWidth() default 15;
    
    /**
     * @Description: 如果是日期类型的字段，那么这个配置就会被采用，如果格式出错，那么就报错，excel导出的代码中不检查该项填写是否合法
     * @return
     *
     * @History
     *     1. 2014-12-24 linwb 创建方法
     */
    String datePattern() default "yyyy-MM-dd";
}
