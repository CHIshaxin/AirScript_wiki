---
title: "Item 方法"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Axes/%E6%96%B9%E6%B3%95/Item%20%E6%96%B9%E6%B3%95.html
---
# Item (方法)

从 **[Axes](./../../Axes/Axes 对象.html)** 集合中返回一个 **[Axis](./../../Axis/Axis 对象.html)** 对象。

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| Type | Kso.XlAxisType | 必填 | 坐标轴类型。 |
| AxisGroup | Default<Kso.XlAxisGroup, 1> | 可选 | 坐标轴。 |

## 示例

```javascript
/*本示例为 Chart1 中分类轴设置标题文本。*/
function test() {
    let axis = Application.Charts.Item("Chart1").ChartObjects(1).Chart.Axes().Item(xlCategory)
    axis.HasTitle = true
    axis.AxisTitle.Caption = "1994"
}
```

```javascript
/*本示例将工作表 Sheet1 中第一张图表的数值轴的主要刻度线设置为跨轴，并将其颜色设置为红色。*/
function test() {
    let axes = Application.Sheets.Item("Sheet1").ChartObjects(1).Chart.Axes()
    axes.Item(xlValue).MajorTickMark = xlTickMarkCross
    axes.Item(xlValue).Border.ColorIndex = 3
}
```
