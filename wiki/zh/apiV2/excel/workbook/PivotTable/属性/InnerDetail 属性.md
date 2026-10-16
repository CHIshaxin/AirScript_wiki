---
title: "InnerDetail (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/PivotTable/%E5%B1%9E%E6%80%A7/InnerDetail%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: B4hoRYJB
---

# InnerDetail (属性)

当最内部行或列字段的 **ShowDetail** 属性设为 **True** 时，返回或设置这些详细数据的字段名称。**String** 类型，可读写。

## 说明

该属性不可用于 OLAP（OLAP：为查询和报表（而不是处理事务）而进行了优化的数据库技术。OLAP 数据是按分级结构组织的，它存储在多维数据集而不是表中。） 数据源。

## 示例

```javascript
/*本示例当最内部行或列字段的 ShowDetail 属性设为 True 时，显示这些详细数据的字段名称。*/
function test() {
    let pvtTable = Application.Worksheets.Item("Sheet1").Range("F1").PivotTable
    //以下代码若报错 则先为InnerDetail设置字段名
    console.log(pvtTable.InnerDetail)
}
```

```javascript
/*本示例为数据透视表设置字段名 z，并显示*/
function test() {
    ActiveSheet.PivotTables(1).InnerDetail = "z"
    console.log(ActiveSheet.PivotTables(1).InnerDetail)
}
```
