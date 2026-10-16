---
title: "Caption (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/ChartTitle/%E5%B1%9E%E6%80%A7/Caption%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: BFdsZbAA
---

# Caption (属性)

返回或设置表示图表标题文本的 **String** 值。

## 示例

```javascript
/*本示例将图表工作表 Chart1 的标题设置为可见，并将标题文本设置为“成绩表”。*/
function test() {
    let chart = Application.Charts.Item("Chart1").ChartObjects(1).Chart
    chart.HasTitle = true
    chart.ChartTitle.Caption = "成绩表"
}
```

```javascript
/*本示例显示工作表 Sheet1 中第一个图表的标题文本是否为“信息表”。*/
function test() {
    let charttitle = Application.Sheets.Item("Sheet1").ChartObjects(1).Chart.ChartTitle
    console.log(charttitle.Caption == "信息表")
}
```
