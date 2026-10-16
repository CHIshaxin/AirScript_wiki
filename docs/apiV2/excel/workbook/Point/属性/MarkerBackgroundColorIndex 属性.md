---
title: "MarkerBackgroundColorIndex 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Point/%E5%B1%9E%E6%80%A7/MarkerBackgroundColorIndex%20%E5%B1%9E%E6%80%A7.html
---
# MarkerBackgroundColorIndex (属性)

返回或设置数据标志的背景色，表示为当前调色板中的索引或下列 **XlColorIndex** 常量之一：**xlColorIndexAutomatic** 或 **xlColorIndexNone**。仅适用于折线图、散点图和雷达图。**Long** 类型，可读写。

## 示例

```javascript
/*此示例为 Chart1 中第一个数据系列的第二个数据点的标记设置背景色和前景色。*/
function test() {
    let point = Application.Charts.Item("Chart1").ChartObjects(1).Chart.SeriesCollection(1).Points(2)
    point.MarkerBackgroundColorIndex = 4  //green
    point.MarkerForegroundColorIndex = 3  //red
}
```

```javascript
/*此示例判断Chart1 中第三个数据系列的第二个数据点标记的背景色如果为无颜色，则其背景色修改为蓝色，反之则其背景色修改为自动配色。*/
function test() {
    let point = Application.Charts.Item("Chart1").ChartObjects(1).Chart.SeriesCollection(3).Points(2)
    if (point.MarkerBackgroundColorIndex == xlColorIndexNone) {
        point.MarkerBackgroundColorIndex = 5
        console.log("数据点标记的背景色为无颜色，已修改为蓝色")
    } else {
        point.MarkerBackgroundColorIndex = xlColorIndexAutomatic
        console.log("数据点标记的背景色不为无颜色，已修改为自动配色")
    }
}
```
