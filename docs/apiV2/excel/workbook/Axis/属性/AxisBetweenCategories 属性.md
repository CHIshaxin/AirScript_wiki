---
title: "AxisBetweenCategories 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Axis/%E5%B1%9E%E6%80%A7/AxisBetweenCategories%20%E5%B1%9E%E6%80%A7.html
---
# AxisBetweenCategories (属性)

如果数值轴与分类坐标轴相交于分类之间，则该属性值为 **True**。**Boolean** 类型，可读写。

## 说明

该属性仅适用于分类轴，而不适用于三维图表。

## 示例

```javascript
/*本示例设置图表工作表 Chart1 的数值轴与分类坐标轴相交于两个分类之间。*/
function test() {
    let axis = Application.Charts.Item("Chart1").ChartObjects(1).Chart.Axes(xlCategory)
    axis.AxisBetweenCategories = true
}
```

```javascript
/*本示例设置工作表 Sheet1 中第一张图表的数值轴与分类坐标轴不相交于两个分类之间，并为该分类轴添加坐标轴标签。*/
function test() {
    let axis = Application.Sheets.Item("Sheet1").ChartObjects(1).Chart.Axes(xlCategory)
    axis.AxisBetweenCategories = false
    axis.HasTitle = true
    axis.AxisTitle.Text = "2023"
}
```
