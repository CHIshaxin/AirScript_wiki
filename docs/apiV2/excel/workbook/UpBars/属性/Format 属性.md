---
title: "Format 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/UpBars/%E5%B1%9E%E6%80%A7/Format%20%E5%B1%9E%E6%80%A7.html
---
# Format (属性)

返回 **[ChartFormat](./../../ChartFormat/ChartFormat 对象.html)** 对象。只读。

## 示例

```javascript
/*本示例将工作表 Sheet1 中第一个图表的第一个图表组的涨柱线的前景色设置为黄色。*/
function test() {
    let upbars = Application.Worksheets.Item("Sheet1").ChartObjects(1).Chart.ChartGroups(1).UpBars
    upbars.Format.Fill.ForeColor.RGB = RGB(255, 255, 0)
}
```

```javascript
/*本示例将图表工作表 Chart1 中图表的第一个图表组的涨柱线的线条设置为蓝色的长划线。*/
function test() {
    let upbars = Application.Charts.Item("Chart1").ChartObjects(1).Chart.ChartGroups(1).UpBars
    upbars.Format.Line.DashStyle = msoLineLongDash
    upbars.Format.Line.ForeColor.RGB = RGB(0, 0, 255)
}
```
