---
title: "Axis 对象"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Axis/Axis%20%E5%AF%B9%E8%B1%A1.html
---
# Axis (对象)

代表图表中的单个坐标轴。

## 说明

**[Axis](./../Axis/Axis 对象.html)** 对象是 **[Axes](./../Axes/Axes 对象.html)** 集合的成员。

使用 **Axes**(*type*, *group* )（其中 *type* 为坐标轴类型，而 *group* 为坐标轴组）可返回单个 **[Axis](./../Axis/Axis 对象.html)** 对象。*Type* 可为以下 **XlAxisType** 常量之一：**xlCategory**、**xlSeries** 或 **xlValue**。*Group* 可为以下 **XlAxisGroup** 常量之一：**xlPrimary** 或 **xlSecondary**。有关详细信息，请参阅 **Axes** 方法。

## 示例

```javascript
/*本示例在名为“Chart1”的图表工作表中设置分类轴的标题文本。*/
function test() {
    let axis = Application.Charts.Item("Chart1").ChartObjects(1).Chart.Axes(xlCategory)
    axis.HasTitle = true
    axis.AxisTitle.Caption = "1994"
}
```

```javascript
/*本示例将工作表 Sheet1 中第一张图表的数值轴的主要刻度线设置为跨轴，并将其颜色设置为红色。*/
function test() {
    let axis = Application.Sheets.Item("Sheet1").ChartObjects(1).Chart.Axes(xlValue)
    axis.MajorTickMark = xlTickMarkCross
    axis.Border.ColorIndex = 3
}
```
