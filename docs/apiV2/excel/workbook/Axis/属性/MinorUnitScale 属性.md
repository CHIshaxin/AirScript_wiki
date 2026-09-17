---
title: "MinorUnitScale 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Axis/%E5%B1%9E%E6%80%A7/MinorUnitScale%20%E5%B1%9E%E6%80%A7.html
---
# MinorUnitScale (属性)

返回或设置当 **CategoryType** 属性设置为 **xlTimeScale** 时分类轴次要单位的刻度值。**XlTimeUnit** 类型，可读写。

## 说明

| **XlTimeUnit** 可为以下 **XlTimeUnit** 常量之一。 |
| --- |
| **xlMonths** |
| **xlDays** |
| **xlYears** |

## 示例

```javascript
/*本示例将图表工作表 Chart1 中图表的分类轴设置为使用时间刻度，并设置主要刻度单位和次要刻度单位。*/
function test() {
    let axis = Application.Charts.Item("Chart1").ChartObjects(1).Chart.Axes(xlCategory)
    axis.CategoryType = xlTimeScale
    axis.MajorUnit = 5
    axis.MajorUnitScale = xlDays
    axis.MinorUnit = 1
    axis.MinorUnitScale = xlDays
}
```

```javascript
/*本示例将工作表 Sheet1 中第一张图表的分类轴设置为使用时间刻度，并设置次要刻度单位和轴标题。*/
function test() {
    let axis = Application.Sheets.Item("Sheet1").ChartObjects(1).Chart.Axes(xlCategory)
    axis.CategoryType = xlTimeScale
    axis.MinorUnitScale = xlYears
    axis.AxisTitle.Caption = "年度报告"
}
```
