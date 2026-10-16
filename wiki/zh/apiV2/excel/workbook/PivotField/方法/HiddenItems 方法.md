---
title: "HiddenItems (方法)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/PivotField/%E6%96%B9%E6%B3%95/HiddenItems%20%E6%96%B9%E6%B3%95.html
updated: 1789700916000
hash: CfRE1tR2
---

# HiddenItems (方法)

返回一个对象，该对象表示单个隐藏数据透视表项（ **[PivotItem](./../../PivotItem/PivotItem 对象.html)** 对象），或表示指定字段中所有隐藏项的集合（ **[PivotItems](./../../PivotItems/PivotItems 对象.html)** 对象）。只读。

## 说明

对于 OLAP（OLAP：为查询和报表（而不是处理事务）而进行了优化的数据库技术。OLAP 数据是按分级结构组织的，它存储在多维数据集而不是表中。） 数据源，该属性总返回一个空集合。

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| Index | any | 可选 | 指定要返回的项的编号或名称（可以是指定多个数据项的数组）。 |

## 示例

```javascript
/*本示例显示活动工作表中数据透视表的字段“score”隐藏的数据透视表项的数量。*/
function test() {
    let pvtField = ActiveSheet.Range("I1").PivotTable.PivotFields("score")
    console.log(pvtField.HiddenItems().Count)
}
```

```javascript
/*本示例显示第一张工作表中第一张数据透视表的字段“name”隐藏的数据透视表项“张八”在字段中的位置。*/
function test() {
    let pvtField = Worksheets.Item(1).PivotTables(1).PivotFields("name")
    console.log(pvtField.HiddenItems("张八").Position)
}
```
