---
title: "MarkerStyle (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Point/%E5%B1%9E%E6%80%A7/MarkerStyle%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: BNXXF-oD
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
/*此示例设置Chart1中第二个系列第三个数据点的样式为正方形。*/
function test() {
    Application.Charts.Item("Chart1").ChartObjects(1).Chart.SeriesCollection(2).Points(3).MarkerStyle = xlMarkerStyleSquare
}
```

```javascript
/*此示例判断Chart1中第四个系列第三个数据点的标志样式如果是自动设置，则其样式修改为菱形，反之其样式修改为三角形。*/
function test() {
    let point = Application.Charts.Item("Chart1").ChartObjects(1).Chart.SeriesCollection(4).Points(3)
    if (point.MarkerStyle == xlMarkerStyleAutomatic) {
        point.MarkerStyle = xlMarkerStyleDiamond
        console.log("数据点标志样式是自动设置，已修改为菱形")
    } else {
        point.MarkerStyle = xlMarkerStyleTriangle
        console.log("数据点标志样式不是自动设置，已修改为三角形")
    }
}
```
