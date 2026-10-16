---
title: "Text 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/ChartTitle/%E5%B1%9E%E6%80%A7/Text%20%E5%B1%9E%E6%80%A7.html
---
# Text (属性)

返回或设置指定对象中的文本。**String** 型，可读写。

## 示例

```javascript
/*此示例设置图表工作表 Chart1 的图表标题文本。*/
function test() {
    let chart = Application.Charts.Item("Chart1").ChartObjects(1).Chart
    chart.HasTitle = true
    chart.ChartTitle.Text = "First Quarter Sales"
}
```

```javascript
/*本示例显示图表工作表 Chart1 的标题文本是否为“成绩表”。*/
function test() {
    let chart = Application.Charts.Item("Chart1").ChartObjects(1).Chart
    console.log(chart.ChartTitle.Text == "成绩表")
}
```
