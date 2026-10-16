---
title: "RightToLeftPlotOrder 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/SparkHorizontalAxis/%E5%B1%9E%E6%80%A7/RightToLeftPlotOrder%20%E5%B1%9E%E6%80%A7.html
---
# RightToLeftPlotOrder (属性)

返回或设置是否按右至左顺序在水平轴上绘制点。可读/写。

## 说明

将 **RightToLeftPlotOrder** 属性设置为 **True** 将按右至左顺序在水平轴上绘制值。

## 返回值

**Boolean**

## 示例

```javascript
/*本示例显示活动工作表上单元格 H1 中第一个迷你图组是否按右至左顺序在水平轴上绘制点。*/
function test() {
    let shAxis = ActiveSheet.Range("H1").SparklineGroups(1).Axes.Horizontal
    console.log(shAxis.RightToLeftPlotOrder)
}
```

```javascript
/*本示例将活动工作表上单元格 H1 中第一个迷你图组的水平轴设置为可见，并将该水平轴设置为按右至左顺序绘制点。*/
function test() {
    let shAxis = ActiveSheet.Range("H1").SparklineGroups(1).Axes.Horizontal
    shAxis.Axis.Visible = true
    shAxis.RightToLeftPlotOrder = true
}
```
