---
title: "MarkerSize (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Point/%E5%B1%9E%E6%80%A7/MarkerSize%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: -g39S7Vv
---

# MarkerSize (属性)

返回或设置数据标志的大小，以磅 （磅：指打印的字符的高度的度量单位。1 磅等于 1/72 英寸，或大约等于 1 厘米的 1/28。）为单位。可以是 2 到 72 之间的一个值。**Long** 型，可读写。

## 示例

```javascript
/*此示例设置Chart1中第一个系列第二个数据点的大小为40磅。*/
function test() {
    Application.Charts.Item("Chart1").ChartObjects(1).Chart.SeriesCollection(1).Points(2).MarkerSize = 40
}
```

```javascript
/*此示例显示Chart1中第三个系列第一个数据点的大小。*/
function test() {
    let point = Application.Charts.Item("Chart1").ChartObjects(1).Chart.SeriesCollection(3).Points(1)
    console.log(`数据点的大小为 ${point.MarkerSize}`)
}
```
