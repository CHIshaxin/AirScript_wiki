---
title: "Format (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Axis/%E5%B1%9E%E6%80%A7/Format%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: pCHaz7j-
---

# Format (属性)

返回 ChartFormat 对象。只读。

## 示例

```javascript
/*本示例将工作表 Sheet1 中第一张图表的数值轴的背景填充色设置为蓝色。*/
function test() {
    let axis = Application.Sheets.Item("Sheet1").ChartObjects(1).Chart.Axes(xlValue)
    axis.Format.Fill.BackColor.RGB = RGB(0, 255, 255)
}
```

```javascript
/*本示例将图表工作表 Chart 中图表的类别轴的线条宽度设置为 3 。*/
function test() {
    let axis = Application.Charts.Item("Chart1").ChartObjects(1).Chart.Axes(xlCategory)
    axis.Format.Line.Weight = 3
}
```
