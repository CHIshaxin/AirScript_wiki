---
title: "TotalLevels (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/PivotField/%E5%B1%9E%E6%80%A7/TotalLevels%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: YikeNUUW
---

# TotalLevels (属性)

返回当前字段组中的字段总数。如果字段没有分组或数据源基于的是 OLAP（OLAP：为查询和报表（而不是处理事务）而进行了优化的数据库技术。OLAP 数据是按分级结构组织的，它存储在多维数据集而不是表中。），则 **TotalLevels** 将返回值 1。**Long**类型，只读。

## 说明

同一分组内的所有字段具有相同的 **TotalLevels** 值。

## 示例

```javascript
/*本示例显示包含活动单元格的字段组中字段的总数。*/
function test() {
    Worksheets.Item("Sheet1").Activate()
    console.log("This group has " + ActiveCell.PivotField.TotalLevels + " levels")
}
```

```javascript
/*本示例显示活动工作表第一张数据透视表的字段“name”中字段组的字段总数。*/
function test() {
    console.log(ActiveSheet.PivotTables(1).PivotFields("name").TotalLevels)
}
```
