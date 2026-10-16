---
title: "FormulaR1C1 (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/DataLabel/%E5%B1%9E%E6%80%A7/FormulaR1C1%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: ByMEayDl
---

# FormulaR1C1 (属性)

获取或设置一个 **String** 值，该值以英语使用 R1C1 样式表示法来表示对象的公式。可读/写。

## 示例

```javascript
/*此示例将图表工作表 Chart1 中图表的第二个数据序列的第三个数据标签设置为工作表 Sheet1 中单元格 A2 的内容。*/
function test() {
    let datalabel = Application.Charts.Item("Chart1").ChartObjects(1).Chart.SeriesCollection(2).DataLabels(3)
    datalabel.FormulaR1C1 = "='Sheet1'!R2C1"
}
```

```javascript
/*此示例将工作表 Sheet1 中第一个图表的第二个数据序列的第二个数据标签设置为该工作表中单元格 C2 的内容。*/
function test() {
    let datalabel = Application.Sheets.Item("Sheet1").ChartObjects(1).Chart.SeriesCollection(2).DataLabels(2)
    datalabel.FormulaR1C1 = "='Sheet1'!R2C3"
}
```
