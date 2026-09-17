---
title: "SparkAxes 对象"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/SparkAxes/SparkAxes%20%E5%AF%B9%E8%B1%A1.html
---
# SparkAxes (对象)

表示一组迷你图的水平轴和垂直轴的设置。

## 说明

使用 **[SparklineGroup](./../SparklineGroup/SparklineGroup 对象.html)** 对象的 **Axes** 属性可返回该组迷你图的 **[SparkAxes](./../SparkAxes/SparkAxes 对象.html)** 对象。

## 示例

```javascript
/*本示例显示活动工作表上单元格 H1 中第一个迷你图组是否按右至左顺序在水平轴上绘制点。*/
function test() {
    let shAxis = ActiveSheet.Range("H1").SparklineGroups(1).Axes.Horizontal
    console.log(shAxis.RightToLeftPlotOrder)
}
```

```javascript
/*本示例设置活动工作表上单元格 H1 中第一个迷你图组的水平轴的颜色。*/
function test() {
    let shAxis = ActiveSheet.Range("H1").SparklineGroups(1).Axes.Horizontal
    shAxis.Axis.Color.ColorIndex = 3
}
```
