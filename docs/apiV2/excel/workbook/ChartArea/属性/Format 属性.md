---
title: "Format 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/ChartArea/%E5%B1%9E%E6%80%A7/Format%20%E5%B1%9E%E6%80%A7.html
---
# Format (属性)

返回 **[ChartFormat](./../../ChartFormat/ChartFormat 对象.html)** 对象。只读。

## 说明

**[ChartFormat](./../../ChartFormat/ChartFormat 对象.html)** 对象包含图表区的线条、填充和效果格式。

## 示例

```javascript
/*此示例将图表工作表 Chart1 的前景填充色设置为绿色。*/
function test() {
    let chartarea = Application.Charts.Item("Chart1").ChartObjects(1).Chart.ChartArea
    chartarea.Format.Fill.ForeColor.RGB = RGB(0, 255, 0)
}
```

```javascript
/*此示例将工作表 Sheet1 中第一个内嵌图表的线条设置为蓝色实线。*/
function test() {
    let chartarea = Application.Worksheets.Item("Sheet1").ChartObjects(1).Chart.ChartArea
    chartarea.Format.Line.DashStyle = msoLineSolid
    chartarea.Format.Line.ForeColor.RGB = RGB(0, 0, 255)
}
```
