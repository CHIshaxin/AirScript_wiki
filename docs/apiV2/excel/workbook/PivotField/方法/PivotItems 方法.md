---
title: "PivotItems 方法"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/PivotField/%E6%96%B9%E6%B3%95/PivotItems%20%E6%96%B9%E6%B3%95.html
---
# PivotItems (方法)

返回一个对象，该对象表示指定字段中的单个数据透视表项（ **[PivotItem](./../../PivotItem/PivotItem 对象.html)** 对象）或所有可见项和隐藏项的集合（ **[PivotItems](./../../PivotItems/PivotItems 对象.html)** 对象）。只读。

## 说明

对于 OLAP（OLAP：为查询和报表（而不是处理事务）而进行了优化的数据库技术。OLAP 数据是按分级结构组织的，它存储在多维数据集而不是表中。） 数据源，该集合是根据唯一的名称（由 **SourceName** 属性返回的名称），而不是根据显示的名称建立索引。

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| Index | any | 可选 | 要返回的项的名称或编号。 |

## 返回值

Variant

## 示例

```javascript
/*本示例将“score”字段中所有项的名称添加到新工作表的列表中。*/
function test() {
    let newSheet = Application.Worksheets.Add()
    newSheet.Activate()
    let pvtTable = Worksheets.Item("Sheet1").Range("I1").PivotTable
    let rw = 0
    for (let i = 1; i <= pvtTable.PivotFields("score").PivotItems().Count; i++) {
        rw++
        newSheet.Cells.Item(rw, 1).Value2 = pvtTable.PivotFields("score").PivotItems(i).Name
    }
}
```

```javascript
/*本示例显示工作表 Sheet1 中第一张数据透视表的字段“name”的第一项的标签文本。*/
function test() {
    let pvtField = Worksheets.Item("Sheet1").PivotTables(1).PivotFields("name")
    console.log(pvtField.PivotItems().Item(1).Caption)
}
```
