---
title: "LeaderLines (对象)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/LeaderLines/LeaderLines%20%E5%AF%B9%E8%B1%A1.html
updated: 1789700916000
hash: Bnju_S6h
---

# LeaderLines (对象)

代表图表的引导线。引导线将数据标签连接到数据点。

## 说明

该对象不是一个集合；没有表示单个引导线的对象。

此对象只适用于饼图。

使用 **LeaderLines** 属性可返回 **[LeaderLines](./../LeaderLines/LeaderLines 对象.html)** 对象。

## 示例

```javascript
/*下例图表一上的数据系列一添加数据标签和蓝色引导线。如果看不到引导线，则该示例代码将失败。在这种情况下，可从饼图中手动拖出其中一个数据标签以便显示引导线。*/
function test() {
    let chartObjects2 = Application.Worksheets.Item(1).ChartObjects(1).Chart.SeriesCollection(1)
    chartObjects2.HasDataLabels = true
    chartObjects2.DataLabels.Position = xlLabelPositionBestFit
    chartObjects2.HasLeaderLines = true
    chartObjects2.LeaderLines.Border.ColorIndex = 5
}
```

```javascript
/*此示例显示图表工作表 Chart1 上的第三个数据系列的引导线粗细。*/
function test() {
    let series3 = Application.Charts.Item("Chart1").ChartObjects(1).Chart.SeriesCollection(3)
    console.log(series3.LeaderLines.Border.Weight)
}
```
