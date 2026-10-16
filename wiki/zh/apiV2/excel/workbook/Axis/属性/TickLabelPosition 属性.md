---
title: "TickLabelPosition (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Axis/%E5%B1%9E%E6%80%A7/TickLabelPosition%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: Eu51dp5z
---

# TickLabelPosition (属性)

说明指定坐标轴上刻度线标志的位置。**XlTickLabelPosition** 类型，可读写。

## 说明

| **XlTickLabelPosition** 可为以下 **XlTickLabelPosition** 常量之一。 |
|---|
| **xlTickLabelPositionLow** |
| **xlTickLabelPositionNone** |
| **xlTickLabelPositionHigh** |
| **xlTickLabelPositionNextToAxis** |

## 示例

```javascript
/*本示例将图表工作表 Chart1 中分类轴的刻度线标签设置为高位（在图表之上）。*/
function test() {
    let axis = Application.Charts.Item("Chart1").ChartObjects(1).Chart.Axes(xlCategory)
    axis.TickLabelPosition = xlTickLabelPositionHigh
}
```

```javascript
/*本示例关闭工作表 Sheet1 中第一张图表的数值轴的刻度线。*/
function test() {
    let axis = Application.Sheets.Item("Sheet1").ChartObjects(1).Chart.Axes(xlValue)
    axis.TickLabelPosition = xlTickLabelPositionNone
}
```
