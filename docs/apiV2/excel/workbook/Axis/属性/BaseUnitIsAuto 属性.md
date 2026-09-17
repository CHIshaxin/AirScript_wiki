---
title: "BaseUnitIsAuto 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Axis/%E5%B1%9E%E6%80%A7/BaseUnitIsAuto%20%E5%B1%9E%E6%80%A7.html
---
# BaseUnitIsAuto (属性)

如果由 ET 为指定的分类轴选取适当的基本单位，则该值为 **True**。默认值为 **True**。**Boolean** 类型，可读写。

## 说明

不能对数值轴设置本属性。

## 示例

```javascript
/*本示例使工作表 Sheet1 上第一个图表中的分类轴使用时间刻度，并使用自动选取的基本单位。*/
function test() {
    let axis = Application.Worksheets.Item("Sheet1").ChartObjects(1).Chart.Axes(xlCategory)
    axis.CategoryType = xlTimeScale
    axis.BaseUnitIsAuto = true
}
```

```javascript
/*本示例使图表工作表 Chart1 中图表的分类轴使用时间刻度，并显示是否自动选取适当的基本单位。*/
function test() {
    let axis = Application.Charts.Item("Chart1").ChartObjects(1).Chart.Axes(xlCategory)
    axis.CategoryType = xlTimeScale
    console.log(axis.BaseUnitIsAuto)
}
```
