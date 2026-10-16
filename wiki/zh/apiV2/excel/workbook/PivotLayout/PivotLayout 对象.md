---
title: "PivotLayout (对象)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/PivotLayout/PivotLayout%20%E5%AF%B9%E8%B1%A1.html
updated: 1789700916000
hash: B8PLjZgH
---

# PivotLayout (对象)

代表数据透视图报表中字段的位置。

## 说明

使用 **PivotLayout** 属性可返回一个 **[PivotLayout](./../PivotLayout/PivotLayout 对象.html)** 对象。

## 示例

```javascript
/*本示例创建在图表工作表 Chart1 上数据透视图报表中所使用的数据透视表字段名称的列表。*/
function test() {
    let newSheet = Worksheets.Add()
    let intRow = 1
    for (let i = 1; i <= Charts.Item("Chart1").PivotLayout.PivotFields().Count; i++) {
        newSheet.Cells.Item(intRow, 1).Value2 = Charts.Item("Chart1").PivotLayout.PivotFields(i).Caption
        intRow++
    }
}
```

```javascript
/*此示例显示工作表 Sheet1 上第一张数据透视图对应的数据透视表的第三个字段的名称。*/
function test() {
    let pvtTable = Worksheets.Item("Sheet1").ChartObjects(1).Chart.PivotLayout.PivotTable
    console.log(pvtTable.PivotFields(3).Name)
}
```
