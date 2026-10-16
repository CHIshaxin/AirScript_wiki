---
title: "Format (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/DropLines/%E5%B1%9E%E6%80%A7/Format%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: hPVC_Hak
---

# Format (属性)

返回 **[ChartFormat](./../../ChartFormat/ChartFormat 对象.html)** 对象。只读。

## 示例

```javascript
/*本示例将图表工作表 Chart1 中图表的第一个图表组的垂直线的线条样式设置为点划线。*/
function test() {
    let droplines = Application.Charts.Item("Chart1").ChartObjects(1).Chart.ChartGroups(1).DropLines
    droplines.Format.Line.DashStyle = msoLineDashDot
}
```

```javascript
/*本示例将工作表 Sheet1 中第一个图表的第一个图表组的垂直线的线条颜色设置为红色。*/
function test() {
    let droplines = Application.Worksheets.Item("Sheet1").ChartObjects(1).Chart.ChartGroups(1).DropLines
    droplines.Format.Line.ForeColor.RGB = RGB(255, 0, 0)
}
```
