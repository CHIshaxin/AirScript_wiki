---
title: "Formula 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/ChartTitle/%E5%B1%9E%E6%80%A7/Formula%20%E5%B1%9E%E6%80%A7.html
---
# Formula (属性)

获取或设置一个 **String** 值，该值以英语使用 A1 样式表示法来表示对象的公式。可读/写。

## 示例

```javascript
/*本示例设置图表工作表 Chart1 中图表的标题为工作表 Sheet1 中 A1 单元格的内容。*/
function test() {
    let charttitle = Application.Charts.Item("Chart1").ChartObjects(1).Chart.ChartTitle
    charttitle.Formula = "=Sheet1!$A$1"
}
```

```javascript
/*本示例将工作表 Sheet1 中第一张图表的标题设置为该表 B5 单元格的内容。*/
function test() {
    let charttitle = Application.Sheets.Item("Sheet1").ChartObjects(1).Chart.ChartTitle
    charttitle.Formula = "=Sheet1!$B$5"
}
```
