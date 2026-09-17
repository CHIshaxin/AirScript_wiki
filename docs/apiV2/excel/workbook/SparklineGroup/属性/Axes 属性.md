---
title: "Axes 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/SparklineGroup/%E5%B1%9E%E6%80%A7/Axes%20%E5%B1%9E%E6%80%A7.html
---
# Axes (属性)

返回关联的 **[SparkAxes](./../../SparkAxes/SparkAxes 对象.html)** 对象。只读。

## 返回值

SparkAxes

## 示例

```javascript
/*本示例显示单元格 I1 中第一个迷你图组垂直轴的MinScaleType属性是否为xlSparkScaleGroup。*/
function test() {
    let sparklineGroup = Range("I1").SparklineGroups.Item(1)
    console.log(sparklineGroup.Axes.Vertical.MinScaleType == xlSparkScaleGroup)
}
```

```javascript
/*本示例将活动工作表上单元格 I1 中第一个迷你图组设置为按右至左顺序在水平轴上绘制点。*/
function test() {
    let sparklineGroup = ActiveSheet.Range("I1").SparklineGroups(1)
    sparklineGroup.Axes.Horizontal.RightToLeftPlotOrder = true
}
```
