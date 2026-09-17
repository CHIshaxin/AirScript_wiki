---
title: "MajorUnitScale 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Axis/%E5%B1%9E%E6%80%A7/MajorUnitScale%20%E5%B1%9E%E6%80%A7.html
---
# MajorUnitScale (属性)

当 **CategoryType** 属性设置为 **xlTimeScale** 时，返回或设置分类轴主要单位的刻度值。**XlTimeUnit** 类型，可读写。

## 说明

| **XlTimeUnit** 可为以下 **XlTimeUnit** 常量之一。 |
| --- |
| **xlMonths** |
| **xlDays** |
| **xlYears** |

## 示例

```javascript
/*本示例将图表工作表 Chart1 中分类轴设置为使用时间刻度，并设置主要刻度单位和次要刻度单位。*/
function test() {
    let axis = Application.Charts.Item("Chart1").ChartObjects(1).Chart.Axes().Item(xlCategory)
    axis.CategoryType = xlTimeScale
    axis.MajorUnit = 5
    axis.MajorUnitScale = xlDays
    axis.MinorUnit = 1
    axis.MinorUnitScale = xlDays
}
```

```javascript
/*本示例将分类轴设置为使用时间刻度，并设置主要刻度单位和次要刻度单位。*/
function test() {
    let axis = Application.Sheets.Item("Sheet1").ChartObjects(1).Chart.Axes().Item(xlCategory)
    axis.CategoryType = xlTimeScale
    axis.MajorUnitScale = xlMonths
}
```
