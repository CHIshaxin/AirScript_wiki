---
title: "Points 对象"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Points/Points%20%E5%AF%B9%E8%B1%A1.html
---
# Points (对象)

图表中指定的系列内所有 **[Point](./../Point/Point 对象.html)** 对象的集合。

## 说明

使用 **Points**(index)（其中 index 是数据点索引号）可返回一个 **[Point](./../Point/Point 对象.html)** 对象。系列中的数据点从左至右编号。`Points(1)` 是最左边的数据点，而 `Points(Points.Count)` 是最右边的数据点。

## 示例

```javascript
/*下例给工作表一上嵌入式图表一中系列一上的最后一个数据点添加数据标签。*/
function test() {
    let points = Worksheets.Item(1).ChartObjects(1).Chart.SeriesCollection(1).Points()
    points.Item(points.Count).ApplyDataLabels(xlShowValue)
}
```

```javascript
/*下例为工作表一上嵌入式图表一中系列一中的第三个数据点设置标志样式。指定的系列必须是 2D 线、散点或雷达系列。*/
function test() {
    Worksheets.Item(1).ChartObjects(1).Chart.SeriesCollection(1).Points(3).MarkerStyle = xlDiamond
}
```
