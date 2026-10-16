---
title: "MinorTickMark (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Axis/%E5%B1%9E%E6%80%A7/MinorTickMark%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: eTsY_8jU
---

# MinorTickMark (属性)

返回或设置指定坐标轴的次要刻度线类型。**XlTickMark** 类型，可读写。

## 说明

| **XlTickMark** 可为以下 **XlTickMark** 常量之一。 |
| --- |
| **xlTickMarkInside** |
| **xlTickMarkOutside** |
| **xlTickMarkCross** |
| **xlTickMarkNone** |

## 示例

```javascript
/*本示例为图表工作表 Chart1 数值轴内侧设置次要刻度线。*/
function test() {
    let axis = Application.Charts.Item("Chart1").ChartObjects(1).Chart.Axes(xlValue)
    axis.MinorTickMark = xlTickMarkInside
}
```

```javascript
/*本示例显示工作表 Sheet1 中第一张图表的分类轴的次要刻度线是否在轴外。*/
function test() {
    let axis = Application.Sheets.Item("Sheet1").ChartObjects(1).Chart.Axes(xlCategory)
    console.log(axis.MinorTickMark == xlTickMarkOutside)
}
```
