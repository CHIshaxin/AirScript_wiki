---
title: "Crosses 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Axis/%E5%B1%9E%E6%80%A7/Crosses%20%E5%B1%9E%E6%80%A7.html
---
# Crosses (属性)

返回或设置指定坐标轴与其他坐标轴相交的点。**Long** 类型，可读写。

## 说明

可以是下表所列的 **XlAxisCrosses** 常量之一。

| 常量 | 含义 |
| --- | --- |
| **xlAxisCrossesAutomatic** | 由 ET 设置坐标轴交点。 |
| **xlMinimum** | 坐标轴在最小值处相交。 |
| **xlMaximum** | 坐标轴在最大值处相交。 |
| **xlAxisCrossesCustom** | **CrossesAt** 属性指定坐标轴交点。 |

该属性对雷达图无效。对于三维图表，该属性只能应用于数值轴并指示由分类轴所定义的平面与数值轴相交的位置。

本属性既可用于分类轴，也可用于数值轴。在分类轴上，xlMinimum 使数值轴在第一个分类处与之相交，而 xlMaximum 使数值轴在最后一个分类处与之相交。

请注意，根据坐标轴的不同，xlMinimum 和 xlMaximum 可能有不同的含义。

## 示例

```javascript
/*本示例设置图表工作表 Chart1 的数值轴与分类坐标轴的交点在最大的 x 值处。*/
function test() {
    let axis = Application.Charts.Item("Chart1").ChartObjects(1).Chart.Axes(xlCategory)
    axis.Crosses = xlMaximum
}
```

```javascript
/*本示例使工作表 Sheet1 中第一个图表的数值轴与分类轴在该分类轴的相交位置由 ET 设置，并通知用户。*/
function test() {
    let axis = Application.Worksheets.Item("Sheet1").ChartObjects(1).Chart.Axes(xlCategory)
    axis.Crosses = xlAxisCrossesAutomatic
    console.log("数值轴与分类轴在分类轴的相交位置由 ET 设置")
}
```
