---
title: "Border (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Axis/%E5%B1%9E%E6%80%A7/Border%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: DLpWl0OX
---

# Border (属性)

返回一个 **[Border](./../../Border/Border 对象.html)** 对象，它代表对象的边框。

## 示例

```javascript
/*本示例将工作表 Sheet1 上第一个图表中的分类轴设置为绿色。*/
function test() {
    let axis = Application.Worksheets.Item("Sheet1").ChartObjects(1).Chart.Axes(xlCategory)
    axis.Border.Color = RGB(0, 255, 0)
}
```

```javascript
/*本示例将图表工作表 Chart1 中图表的数值轴设置为粗边框，并将其颜色设置为红色。*/
function test() {
    let axis = Application.Charts.Item("Chart1").ChartObjects(1).Chart.Axes(xlValue)
    axis.Border.Weight = xlThick
    axis.Border.ColorIndex  = 3
}
```
