---
title: "MarkerForegroundColor 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Point/%E5%B1%9E%E6%80%A7/MarkerForegroundColor%20%E5%B1%9E%E6%80%A7.html
---
# MarkerForegroundColor (属性)

将数据标志的背景色设置为 RGB 值或返回对应的颜色索引值。仅适用于折线图、散点图和雷达图。**Long** 型，可读写。

## 示例

```javascript
/*此示例为 Chart1 中第一个数据系列的第二个数据点的标记设置背景色和前景色。*/
function test() {
    let point = Application.Charts.Item("Chart1").ChartObjects(1).Chart.SeriesCollection(1).Points(2)
    point.MarkerBackgroundColor = RGB(0, 255, 0)  //green fill
    point.MarkerForegroundColor = RGB(255, 0, 0)  //red border
}
```

```javascript
/*此示例将 Chart1 中第三个数据系列的偶数位的数据点的前景色设置为白色。*/
function test() {
    let points = Application.Charts.Item("Chart1").ChartObjects(1).Chart.SeriesCollection(3).Points()
    for (let i = 2; i <= points.Count; i += 2) {
        points.Item(i).MarkerForegroundColor = RGB(255, 255, 255)
    }
}
```
