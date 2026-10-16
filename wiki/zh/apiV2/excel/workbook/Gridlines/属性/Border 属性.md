---
title: "Border (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Gridlines/%E5%B1%9E%E6%80%A7/Border%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: DSNVmTrp
---

# Border (属性)

返回一个 **[Border](./../../Border/Border 对象.html)** 对象，它代表对象的边框。

## 示例

```javascript
/*本示例打开工作表 Sheet1 中第一个图表类别轴的次要网格线，并将网格线边框设置为绿色。*/
function test() {
    let axis = Worksheets.Item("Sheet1").ChartObjects(1).Chart.Axes(xlCategory)
    axis.HasMinorGridlines = true
    axis.MinorGridlines.Border.ColorIndex = 4
}
```

```javascript
/*本示例显示图表工作表 Chart1 数值轴的主要网格线的边框粗细是否为细线 。*/
function test() {
    let axis = Application.Charts.Item("Chart1").ChartObjects(1).Chart.Axes(xlValue)
    console.log(axis.MajorGridlines.Border.Weight == xlHairline)
}
```
