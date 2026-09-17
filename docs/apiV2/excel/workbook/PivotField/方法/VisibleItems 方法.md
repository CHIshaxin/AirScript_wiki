---
title: "VisibleItems 方法"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/PivotField/%E6%96%B9%E6%B3%95/VisibleItems%20%E6%96%B9%E6%B3%95.html
---
# VisibleItems (方法)

返回一个对象，该对象表示指定字段中的单个可见数据透视表项（ **[PivotItem](./../../PivotItem/PivotItem 对象.html)** 对象），或指定字段中的所有可见项的集合（ **[PivotItems](./../../PivotItems/PivotItems 对象.html)** 对象）。只读。

## 说明

对于 OLAP（OLAP：为查询和报表（而不是处理事务）而进行了优化的数据库技术。OLAP 数据是按分级结构组织的，它存储在多维数据集而不是表中。） 数据源，本属性是只读的，并且总返回 **True**。没有隐藏项。

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| Index | any | 可选 | 要返回的项的编号或名称（可以是指定多项的数组）。 |

## 示例

```javascript
/*本示例向新工作表上的列表中添加“id”字段中所有可见的数据项名称。*/
function test() {
    let nwSheet = Worksheets.Add()
    nwSheet.Activate()
    let pvtTable = Worksheets.Item("Sheet1").Range("I1").PivotTable
    let rw = 0
    for (let i = 1; i <= pvtTable.PivotFields("id").VisibleItems().Count; i++) {
        rw++
        nwSheet.Cells.Item(rw, 1).Value2 = pvtTable.PivotFields("id").VisibleItems().Item(i).Name
    }
}
```

```javascript
/*本示例显示工作表 Sheet1 中第一张数据透视表的字段“score”的第二个可见项的名称。*/
function test() {
    console.log(Worksheets.Item("Sheet1").PivotTables(1).PivotFields("score").VisibleItems(2).Name)
}
```
