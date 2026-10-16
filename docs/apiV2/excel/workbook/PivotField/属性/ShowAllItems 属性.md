---
title: "ShowAllItems 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/PivotField/%E5%B1%9E%E6%80%A7/ShowAllItems%20%E5%B1%9E%E6%80%A7.html
---
# ShowAllItems (属性)

如果显示数据透视表中的所有项目（即使这些项目中不包含汇总数据），则该值为 **True**。默认值为 **False**。**Boolean** 类型，可读写。

## 说明

对于 OLAP（OLAP：为查询和报表（而不是处理事务）而进行了优化的数据库技术。OLAP 数据是按分级结构组织的，它存储在多维数据集而不是表中。） 数据源，该值总是 **False**。

## 示例

```javascript
/*本示例显示第一个工作表上第一个数据透视表中“Month”字段的所有行，包括没有数据的月份。*/
function test() {
    Worksheets.Item(1).PivotTables(1).PivotFields("Month").ShowAllItems = true
}
```

```javascript
/*本示例显示活动工作表中第一张数据透视表的字段“name”是否展示所有项。*/
function test() {
    console.log(ActiveSheet.PivotTables(1).PivotFields("name").ShowAllItems)
}
```
