---
title: "BaseUnit 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Axis/%E5%B1%9E%E6%80%A7/BaseUnit%20%E5%B1%9E%E6%80%A7.html
---
# BaseUnit (属性)

返回或设置指定分类轴的基本单位。**XlTimeUnit** 类型，可读写。

## 说明

如果指定坐标轴的 **CategoryType** 属性设置为 **xlCategoryScale**，则设置该属性不会看到任何效果。不过，将保留设置值，并在 **CategoryType** 属性设置为 **xlTimeScale** 时起作用。

不能对数值轴设置本属性。

## 示例

```javascript
/*本示例使第一张工作表上第一个图表中的分类轴使用时间刻度，并以月为基本单位。*/
function test() {
    let axis = Application.Worksheets.Item(1).ChartObjects(1).Chart.Axes(xlCategory)
    axis.CategoryType = xlTimeScale
    axis.BaseUnit = xlMonths
}
```

```javascript
/*本示例使工作表 Sheet1 上图表 Chart 2 中的分类轴使用时间刻度，并显示是否以天为基本单位。*/
function test() {
    let axis = Application.Worksheets.Item("Sheet1").ChartObjects("Chart 2").Chart.Axes(xlCategory)
    axis.CategoryType = xlTimeScale
    console.log(axis.BaseUnit == xlDays)
}
```
