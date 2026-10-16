---
title: "Horizontal 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/SparkAxes/%E5%B1%9E%E6%80%A7/Horizontal%20%E5%B1%9E%E6%80%A7.html
---
# Horizontal (属性)

返回指定 **[SparkAxes](./../../SparkAxes/SparkAxes 对象.html)** 对象的 **[SparkHorizontalAxis](./../../SparkHorizontalAxis/SparkHorizontalAxis 对象.html)** 对象。只读

## 返回值

**[SparkHorizontalAxis](./../../SparkHorizontalAxis/SparkHorizontalAxis 对象.html)**

## 示例

```javascript
/*本示例将活动工作表上单元格 H1 中第一个迷你图组设置为按右至左顺序在水平轴上绘制点。*/
function test() {
    let sparklineGroup = ActiveSheet.Range("H1").SparklineGroups(1)
    sparklineGroup.Axes.Horizontal.RightToLeftPlotOrder = true
}
```

```javascript
/*本示例显示活动工作表上单元格 H1 中第一个迷你图组的迷你图的水平轴是否基于日期值。*/
function test() {
    let shAxis = ActiveSheet.Range("H1").SparklineGroups(1).Axes.Horizontal
    console.log(shAxis.IsDateAxis)
}
```
