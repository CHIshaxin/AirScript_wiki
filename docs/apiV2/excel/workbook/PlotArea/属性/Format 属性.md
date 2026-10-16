---
title: "Format 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/PlotArea/%E5%B1%9E%E6%80%A7/Format%20%E5%B1%9E%E6%80%A7.html
---
# Format (属性)

返回 **[ChartFormat](./../../ChartFormat/ChartFormat 对象.html)** 对象。只读。

## 示例

```javascript
/*本示例在 Chart1 中的绘图区右侧偏移 5 磅添加阴影。*/
function test() {
    let plotArea = Application.Charts.Item("Chart1").ChartObjects(1).Chart.PlotArea
    plotArea.Format.Shadow.Visible = true
    plotArea.Format.Shadow.OffsetX = 5
}
```

```javascript
/*下例将 Sheet1 上第一个图表的绘图区背景色设为红色。*/
function test() {
    let plotArea = Application.Worksheets.Item("Sheet1").ChartObjects(1).Chart.PlotArea
    plotArea.Format.Fill.BackColor.RGB = RGB(255, 0, 0)
}
```
