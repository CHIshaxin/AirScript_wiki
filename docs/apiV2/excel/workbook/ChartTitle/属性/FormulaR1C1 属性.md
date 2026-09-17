---
title: "FormulaR1C1 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/ChartTitle/%E5%B1%9E%E6%80%A7/FormulaR1C1%20%E5%B1%9E%E6%80%A7.html
---
# FormulaR1C1 (属性)

获取或设置一个 **String** 值，该值以英语使用 R1C1 样式表示法来表示对象的公式。可读/写。

## 示例

```javascript
/*本示例设置图表工作表 Chart1 的标题为工作表 Sheet1 中 B2 单元格的内容。*/
function test() {
    let charttitle = Application.Charts.Item("Chart1").ChartObjects(1).Chart.ChartTitle
    charttitle.FormulaR1C1 = "='Sheet1'!R2C2"
}
```

```javascript
/*本示例设置工作表 Sheet1 上第一张图表的标题为该表 A4 单元格的内容。*/
function test() {
    let charttitle = Application.Sheets.Item("Sheet1").ChartObjects(1).Chart.ChartTitle
    charttitle.FormulaR1C1 = "='Sheet1'!R4C1"
}
```
