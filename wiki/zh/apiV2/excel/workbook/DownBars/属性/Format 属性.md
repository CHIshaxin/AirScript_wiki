---
title: "Format (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/DownBars/%E5%B1%9E%E6%80%A7/Format%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: 7Z3C8wSq
---

# Format (属性)

返回 **[ChartFormat](./../../ChartFormat/ChartFormat 对象.html)** 对象。只读。

## 示例

```javascript
/*本示例将工作表 Sheet1 中第一个图表的第一个图表组的跌柱线的前景色设置为绿色。*/
function test() {
    let downbars = Application.Worksheets.Item("Sheet1").ChartObjects(1).Chart.ChartGroups(1).DownBars
    downbars.Format.Fill.ForeColor.RGB = RGB(0, 255, 0)
}
```

```javascript
/*本示例将图表工作表 Chart1 中图表的第一个图表组的跌柱线的线条设置为红色的点划线。*/
function test() {
    let downbars = Application.Charts.Item("Chart1").ChartObjects(1).Chart.ChartGroups(1).DownBars
    downbars.Format.Line.DashStyle = msoLineDashDot
    downbars.Format.Line.ForeColor.RGB = RGB(255, 0, 0)
}
```
