---
title: "Point (对象)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Point/Point%20%E5%AF%B9%E8%B1%A1.html
updated: 1789700916000
hash: jCPVmrqz
---

# Point (对象)

代表图表系列中的单个数据点。

## 说明

**[Point](./../Point/Point 对象.html)** 对象是 **[Points](./../Points/Points 对象.html)** 集合的成员。 **[Points](./../Points/Points 对象.html)** 集合包含一个系列中所有的数据点。

使用 **Points**(index)（其中 index 是数据点索引号）可返回一个 **[Point](./../Point/Point 对象.html)** 对象。系列中的数据点从左至右编号。`Points(1)` 是最左边的数据点，而 `Points(Points.Count)` 是最右边的数据点。

## 示例

```javascript
/*下例为工作表一上嵌入式图表一中系列一中的第三个数据点设置标志样式。指定的系列必须是 2D 线、散点或雷达系列。*/
function test() {
    Worksheets.Item(1).ChartObjects(1).Chart.SeriesCollection(1).Points(3).MarkerStyle = xlDiamond
}
```

```javascript
/*本示例将 Chart1 中第一个数据系列的第4个数据点设置成有阴影。*/
function test() {
    let series = Application.Charts.Item("Chart1").ChartObjects(1).Chart.SeriesCollection(1)
    series.Points(4).Shadow = true
}
```
