---
title: "PivotFormula 对象"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/PivotFormula/PivotFormula%20%E5%AF%B9%E8%B1%A1.html
---
# PivotFormula (对象)

代表在数据透视表中用于计算的公式。

## 说明

本对象及其相关属性和方法对于 OLAP（OLAP：为查询和报表（而不是处理事务）而进行了优化的数据库技术。OLAP 数据是按分级结构组织的，它存储在多维数据集而不是表中。） 数据源无效，这是因为它不支持计算字段和计算项。

使用 **PivotFormulas** (*index*)（其中 *index* 是公式左侧的公式号或字符串）可返回 **[PivotFormula](./../PivotFormula/PivotFormula 对象.html)** 对象。

## 示例

```javascript
/*本示例将更改公式一（第一张工作表的第一个数据透视表中）的索引号，使其在公式二计算完毕后再进行计算。*/
function test() {
    Worksheets.Item(1).PivotTables(1).PivotFormulas.Item(1).Index = 2
}
```

```javascript
/*本示例删除第一张工作表中数据透视表的第二个公式。*/
function test() {
    let pvtFormula = Worksheets.Item(1).Range("I1").PivotTable.PivotFormulas(2)
    pvtFormula.Delete()
}
```
