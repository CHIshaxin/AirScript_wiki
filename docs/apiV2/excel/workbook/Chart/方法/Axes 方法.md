---
title: "Axes 方法"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Chart/%E6%96%B9%E6%B3%95/Axes%20%E6%96%B9%E6%B3%95.html
---
# Axes (方法)

返回一个代表图表上单个坐标轴或坐标轴集合的某个对象。

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| Type | any | 可选 | 指定要返回的坐标轴。可为以下 XlAxisType 常量之一：xlValue、xlCategory 或 xlSeriesAxis（xlSeriesAxis 仅对三维图表有效）。 |
| AxisGroup | Default<Kso.XlAxisGroup, 1> | 可选 | 指定坐标轴组。如果省略该参数，则使用主坐标轴组。三维图表仅有一个坐标轴组。 |

## 示例

```javascript
/*本示例为图表工作表 Chart1 的分类轴添加轴标签。*/
function test() {
    let axes = Application.Charts.Item("Chart1").Axes(xlCategory)
    axes.HasTitle = true
    axes.AxisTitle.Text = "July Sales"
}
```

```javascript
/*本示例将关闭图表工作表 Chart1 中分类轴的主要网格线。*/
function test() {
    Application.Charts.Item("Chart1").Axes(xlCategory).HasMajorGridlines = false
}
```

```javascript
/*本示例将关闭 Chart1 中所有坐标轴的网格线。*/
function test() {
    for (let i = 1; i <= Application.Charts.Item("Chart1").Axes().Count; i++) {
        Application.Charts.Item("Chart1").Axes().Item(i).HasMajorGridlines = false
        Application.Charts.Item("Chart1").Axes().Item(i).HasMinorGridlines = false
    }
}
```
