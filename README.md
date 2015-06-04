# excelExportor
Excel导出通用接口及实现

##API介绍
1. 对于要导出的实体类，必须用JAVABEAN的形式表示.  
2. 目前不支持的属性类型为：boolean,List,Array  但是可以通过另类的做法，比如：Boolean类型的，那么就采用字符串的形式表示 （如：男/女）

源码： ExcelExportor ,  注解类：ExcelExport , 实体类：ExcelHeader

例子程序：ExcelExportExample

只有用 @ExcelExport 注解标记的字段才会被导出
针对@ExcelExport的说明


+ ExcelExport的colWidth用来指定每一个列的宽度，默认值15.

+ ExcelExport的header当前版本采用字符串表示,以后如果要扩展在进行相应的修改，
	比如要针对不同的地区，同一个字段导出的表头名称又不一样，那么可以通过给header配置一个key，然后具体的表头名称配置到数据库，根据KEY从数据库或缓存中获取对应的header值。

+ 针对数据量太多的时候, 默认每超过 5000 条数据，就在创建一个表格

+ ExcelExport用来创建表格，那么就要有一些表格的特性，我们的这个注解是从Html的table中借鉴了 colspan 和 rowspan这两个属性，然后我们在此基础上再进行扩展，
	对于多表头的配置我们后面会详解，我们这里先混个眼熟，多表头是用逗号分割多个数值来表示的，比如： rowspan="1，2，3" 就表示  从下往上（最接近导出数据的那一行表头开始）数，
	分别为一行，两行，三行 那么总共就有 6 行; 这样的话真正要导出的数据就那个 第7行开始


##总的原则 以及 怎么使用

> 从最靠近数据的那一行开始往上设置,最靠近数据的那一行的colspan永远都是1,然后往上级表头遍历

比如： 对于下面这个表格的设置  colspan 和 rowspan 要这样设置：
<数字 代表列>(colspan="" , rowspan="")

`1(colspan="1" rowspan="3");`
从这一列开始 有 父级表头,那么就需要设置父级表头,用逗号隔开,往上设置colspan和rowspan
 第四行开始就是数据行了，所以最靠近数据行的就是第三行,而第二列开始出现父级表头,所以,从第二列开始设置父级表头

 第二列的直接父级表头是 13 再上一级的父级表头是  14  所以  colspan="第2列的colspan,第13列的colspan,第14列的colspan"

 同理： 因为第2列,第13列,第14列都占据一行 所以:  rowspan="第2列的rowspan,第13列的rowspan,第14列的rowspan"

 因此 计算出第2列的  设置应该是：  (colspan="1,2,5",rowspan="1,1,1)

 注: 如果cospan和rowspan有多个值, 设置header的时候也要设置多个值
	 
	2(colspan="1,2,5" rowspan="1,1,1"); 
	3(colspan="1" rowspan="1");
	
	4(colspan="1,3" rowspan="1,1");
	5(colspan="1" rowspan="1");
	6(colspan="1" rowspan="1");
	
	7(colspan="1,6" rowspan="2,1");
	8(colspan="1,5" rowspan="1,1");
	9(colspan="1" rowspan="1");
	10(colspan="1" rowspan="1");
	11(colspan="1" rowspan="1");
	12(colspan="1" rowspan="1");

	-------------------------------------------------
	| 合	|		14			|		17				|
	-	---------------------------------------------
	| 并	|	13	|	15		|   |		18			|	
	-	---------------------   ---------------------
	| 1	| 2	| 3	| 4	| 5	| 6	| 7	| 8	| 9 |10	|11	|12	|
	-------------------------------------------------


##图示说明
![Excel导出工具使用图解](http://i1.tietuku.com/4abfd84fffcf2c2c.png)