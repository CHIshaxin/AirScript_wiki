---
title: "MarkerStyle 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Series/%E5%B1%9E%E6%80%A7/MarkerStyle%20%E5%B1%9E%E6%80%A7.html
---
# MarkerStyle (属性)

返回或设置折线图、散点图或雷达图中数据点或数据系列的数据标志样式。**XlMarkerStyle** 类型，可读写。

## 说明

| **XlMarkerStyle** 可为以下 **XlMarkerStyle** 常量之一。 |
| --- |
| **xlMarkerStyleAutomatic**。自动设置标记 |
| **xlMarkerStyleCircle**。圆形标记 |
| **xlMarkerStyleDash**。长条形标记 |
| **xlMarkerStyleDiamond**。菱形标记 |
| **xlMarkerStyleDot**。短条形标记 |
| **xlMarkerStyleNone**。无标记 |
| **xlMarkerStylePicture**。图片标记 |
| **xlMarkerStylePlus**。带加号的方形标记 |
| **xlMarkerStyleSquare**。方形标记 |
| **xlMarkerStyleStar**。带星号的方形标记 |
| **xlMarkerStyleTriangle**。三角形标记 |
| **xlMarkerStyleX**。带 X 记号的方形标记 |

## 示例

```javascript
/*本示例设置 Chart1 中第一个数据系列的标记样式。本示例应在二维折线图上运行。*/
function test() {
    Application.Charts.Item("Chart1").ChartObjects(1).Chart.SeriesCollection(1).MarkerStyle = xlMarkerStyleCircle
}
```

```javascript
/*此示例显示第一张工作表上的第一个图表的第三个系列数据标记样式是否为无标记。*/
function test() {
    console.log(Application.Worksheets.Item(1).ChartObjects(1).Chart.SeriesCollection(3).MarkerStyle == xlMarkerStyleNone)
}
```
