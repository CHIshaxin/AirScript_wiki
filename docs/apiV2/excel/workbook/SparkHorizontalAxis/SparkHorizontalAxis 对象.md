---
title: "SparkHorizontalAxis 对象"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/SparkHorizontalAxis/SparkHorizontalAxis%20%E5%AF%B9%E8%B1%A1.html
---
# SparkHorizontalAxis (对象)

表示一组迷你图的水平轴的设置。

## 说明

使用 **[SparkAxes](./../SparkAxes/SparkAxes 对象.html)** 对象的 **Horizontal** 属性可返回迷你图组的 **[SparkHorizontalAxis](./../SparkHorizontalAxis/SparkHorizontalAxis 对象.html)** 对象。水平轴只有在迷你图数据在垂直轴上同时有负值和正值的情况下才会显示。

## 示例

```javascript
/*本示例将活动工作表上单元格 H1 中第一个迷你图组的水平轴设置为可见，并设置该水平轴的颜色。*/
function test() {
    let shAxis = ActiveSheet.Range("H1").SparklineGroups(1).Axes.Horizontal
    shAxis.Axis.Visible = true
    shAxis.Axis.Color.ColorIndex = 3
}
```

```javascript
/*本示例判断活动工作表上单元格 H1 中第一个迷你图组的水平轴绘制点的顺序，并通知用户。*/
function test() {
    let shAxis = ActiveSheet.Range("H1").SparklineGroups(1).Axes.Horizontal
    if (shAxis.RightToLeftPlotOrder) {
        console.log("按右至左顺序在水平轴上绘制点")
    } else {
        console.log("按左至右顺序在水平轴上绘制点")
    }
}
```
