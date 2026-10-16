---
title: "FormulaLocal (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/DataLabel/%E5%B1%9E%E6%80%A7/FormulaLocal%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: CMRAn5kU
---

# FormulaLocal (属性)

获取或设置一个 **String** 值，该值以用户的语言使用 A1 样式表示法来表示对象的公式。可读/写。

## 示例

```javascript
/*此示例将图表工作表 Chart1 中图表的第二个数据序列的第二个数据标签设置为工作表 Sheet1 中单元格 B1 的内容。*/
function test() {
    let datalabel = Application.Charts.Item("Chart1").ChartObjects(1).Chart.SeriesCollection(2).DataLabels(2)
    datalabel.FormulaLocal = "=Sheet1!$B$1"
}
```

```javascript
/*此示例将工作表 Sheet1 中第一个图表的第一个数据序列的第二个数据标签设置为该工作表中单元格 A5 的内容。*/
function test() {
    let datalabel = Application.Sheets.Item("Sheet1").ChartObjects(1).Chart.SeriesCollection(1).DataLabels(2)
    datalabel.FormulaLocal = "=Sheet1!$A$5"
}
```
