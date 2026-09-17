---
title: "CategoryType 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Axis/%E5%B1%9E%E6%80%A7/CategoryType%20%E5%B1%9E%E6%80%A7.html
---
# CategoryType (属性)

返回或设置分类轴类型。**XlCategoryType** 类型，可读写。

## 说明

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
/*本示例显示工作表 Sheet1 中图表 Chart 2 是否为 ET 控制分类轴的类型。*/
function test() {
    let axis = Application.Worksheets.Item("Sheet1").ChartObjects("Chart 2").Chart.Axes(xlCategory)
    console.log(axis.CategoryType == xlAutomaticScale)
}
```
