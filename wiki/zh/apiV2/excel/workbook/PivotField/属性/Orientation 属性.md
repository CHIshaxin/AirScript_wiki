---
title: "Orientation (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/PivotField/%E5%B1%9E%E6%80%A7/Orientation%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: DmHzDlk8
---

# Orientation (属性)

返回或设置一个 **XlPivotFieldOrientation** 值，它代表字段在指定的数据透视表中的位置。

## 说明

对于 OLAP（OLAP：为查询和报表（而不是处理事务）而进行了优化的数据库技术。OLAP 数据是按分级结构组织的，它存储在多维数据集而不是表中。）数据源，当设置某一层中的一个字段的此属性值时，也会同时设置同一层中所有其他字段的方向。维字段只能在数据透视表的行、列和页字段区域中进行定向。度量字段则只能在数据区域中进行定向。将某一层或数据字段设置为 **xlHidden** 时，会将该层或字段从数据透视表中移出。

## 示例

```javascript
/*本示例显示“name”字段的方向。*/
function test() {
    let pvtTable = Worksheets.Item("Sheet1").Range("I1").PivotTable
    let pvtField = pvtTable.PivotFields("name")
    switch (pvtField.Orientation) {
        case xlHidden:
            console.log("Hidden field")
            break
        case xlRowField:
            console.log("Row field")
            break
        case xlColumnField:
            console.log("Column field")
            break
        case xlPageField:
            console.log("Page field")
            break
        case xlDataField:
            console.log("Data field")
    }
}
```

```javascript
/*本示例将活动工作表中第一张数据透视表的字段“score”的在该透视表的位置设置为“xlDataField”。*/
function test() {
    ActiveSheet.PivotTables(1).PivotFields("score").Orientation = xlDataField
}
```
