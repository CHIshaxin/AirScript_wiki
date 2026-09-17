---
title: "Text 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/AxisTitle/%E5%B1%9E%E6%80%A7/Text%20%E5%B1%9E%E6%80%A7.html
---
# Text (属性)

返回或设置指定对象中的文本。**String** 型，可读写。

## 示例

```javascript
/*此示例设置图表工作表 Chart1 中图表的分类轴标题文本。*/
function test() {
    let axis = Application.Charts.Item("Chart1").ChartObjects(1).Chart.Axes(xlCategory)
    axis.HasTitle = true
    axis.AxisTitle.Text = "Month"
}
```

```javascript
/*此示例显示工作表 Sheet1 中第一张图表的数值轴标题文本。*/
function test() {
    let axistitle = Application.Sheets.Item("Sheet1").ChartObjects(1).Chart.Axes(xlValue)
    console.log(axistitle.AxisTitle.Text)
}
```
