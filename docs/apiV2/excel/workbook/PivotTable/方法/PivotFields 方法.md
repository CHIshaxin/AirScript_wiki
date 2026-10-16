---
title: "PivotFields 方法"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/PivotTable/%E6%96%B9%E6%B3%95/PivotFields%20%E6%96%B9%E6%B3%95.html
---
# PivotFields (方法)

返回一个对象，该对象表示数据透视表中单个字段（ **[PivotField](./../../PivotField/PivotField 对象.html)** 对象）或可见及隐藏字段集合（ **[PivotFields](./../../PivotFields/PivotFields 对象.html)** 对象）。只读。

## 说明

对于 OLAP（OLAP：为查询和报表（而不是处理事务）而进行了优化的数据库技术。OLAP 数据是按分级结构组织的，它存储在多维数据集而不是表中。）数据源，不存在隐藏字段，返回的对象或集合反映了当前可见的内容。

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| Index | any | 可选 | 要返回的字段的名称或编号。 |

## 返回值

Object

## 示例

```javascript
/*本示例向新工作表列表添加数据透视表的字段名称。*/
function test() {
    let nwSheet = Worksheets.Add()
    nwSheet.Activate()
    let pvtTable = Worksheets.Item("Sheet1").Range("F1").PivotTable
    let rw = 0
    for (let i = 1; i <= pvtTable.PivotFields().Count; i++) {
        rw++
        nwSheet.Cells.Item(rw, 1).Value2 = pvtTable.PivotFields(i).Name
    }
}
```

```javascript
/*本示例显示活动工作表中第一张数据透视表的字段数量*/
function test() {
    console.log(ActiveSheet.PivotTables(1).PivotFields().Count)
}
```
