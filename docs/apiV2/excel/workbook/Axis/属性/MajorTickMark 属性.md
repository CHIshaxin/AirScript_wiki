---
title: "MajorTickMark 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Axis/%E5%B1%9E%E6%80%A7/MajorTickMark%20%E5%B1%9E%E6%80%A7.html
---
# MajorTickMark (属性)

返回或设置指定坐标轴的主刻度线类型。**XlTickMark** 类型，可读写。

## 说明

| **XlTickMark** 可为以下 **XlTickMark** 常量之一。 |
| --- |
| **xlTickMarkInside** |
| **xlTickMarkOutside** |
| **xlTickMarkCross** |
| **xlTickMarkNone** |

## 示例

```javascript
/*本示例将图表工作表 Chart1 中的数值轴主要刻度线设置为在轴外侧。*/
function test() {
    let axis = Application.Charts.Item("Chart1").ChartObjects(1).Chart.Axes().Item(xlValue)
    axis.MajorTickMark = xlTickMarkOutside
}
```

```javascript
/*本示例将工作表 Sheet1 中第一张图表的分类轴的主要刻度线设置为跨轴，并将其颜色设置为红色。*/
function test() {
    let axis = Application.Sheets.Item("Sheet1").ChartObjects(1).Chart.Axes(xlCategory)
    axis.MajorTickMark = xlTickMarkCross
    axis.Border.ColorIndex = 3
}
```
