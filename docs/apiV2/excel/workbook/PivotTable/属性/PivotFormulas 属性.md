---
title: "PivotFormulas 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/PivotTable/%E5%B1%9E%E6%80%A7/PivotFormulas%20%E5%B1%9E%E6%80%A7.html
---
# PivotFormulas (属性)

返回一个 **[PivotFormulas](./../../PivotFormulas/PivotFormulas 对象.html)** 对象，该对象表示指定数据透视表的公式集合。只读。

## 说明

对于 OLAP（OLAP：为查询和报表（而不是处理事务）而进行了优化的数据库技术。OLAP 数据是按分级结构组织的，它存储在多维数据集而不是表中。） 数据源，该属性将返回一个空集合。

## 示例

```javascript
/*本示例将活动工作表中第一张数据透视表的公式赋值到单元格中。*/
function test() {
    let r = 10
    for (let i = 1; i <= Application.ActiveSheet.PivotTables(1).PivotFormulas.Count; i++) {
        r++
        Cells.Item(r, 1).Value2 = Application.ActiveSheet.PivotTables(1).PivotFormulas.Item(i).Formula
    }
}
```

```javascript
/*本示例显示工作表 Sheet1 中第一张数据透视表的第一个公式的名称*/
function test() {
    let pvtTable = Worksheets.Item("Sheet1").PivotTables(1)
    console.log(pvtTable.PivotFormulas.Item(1).Value)
}
```
