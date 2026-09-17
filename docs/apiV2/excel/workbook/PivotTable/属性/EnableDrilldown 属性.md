---
title: "EnableDrilldown 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/PivotTable/%E5%B1%9E%E6%80%A7/EnableDrilldown%20%E5%B1%9E%E6%80%A7.html
---
# EnableDrilldown (属性)

如果启用“显示明细数据”，则该值为 **True**。默认值为 **True**。**Boolean** 类型，可读写。

## 说明

为数据透视表设置该属性将对该表中的所有字段有效。

对于 OLAP（OLAP：为查询和报表（而不是处理事务）而进行了优化的数据库技术。OLAP 数据是按分级结构组织的，它存储在多维数据集而不是表中。） 数据源，该属性值总是 **True**。

## 示例

```javascript
/*本示例对第一张工作表上第一个数据透视表中的所有字段禁用明细。*/
function test() {
    Application.Worksheets.Item(1).PivotTables("Pivot1").EnableDrilldown = false
}
```

```javascript
/*本示例显示活动工作表中第一张数据透视表 EnableDrilldown 属性的数据类型，并设置为true*/
function test() {
    console.log(typeof ActiveSheet.PivotTables("Pivot1").EnableDrilldown)
    ActiveSheet.PivotTables("Pivot1").EnableDrilldown = true
}
```
