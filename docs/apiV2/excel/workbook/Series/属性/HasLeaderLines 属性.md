---
title: "HasLeaderLines 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Series/%E5%B1%9E%E6%80%A7/HasLeaderLines%20%E5%B1%9E%E6%80%A7.html
---
# HasLeaderLines (属性)

如果数据系列有引导线，则该属性值为 **True**。**Boolean** 类型，可读写。

## 说明

该属性仅应用于饼图。

## 示例

```javascript
/*本示例向饼图上的数据系列 1 添加数据标签和蓝色引导线。如果看不到引导线，则本示例代码将失败。这种情况下，可以手动将数据标签之一拖出饼图以便让引导线显示。*/
function test() {
    let Ser = Application.Worksheets.Item(1).ChartObjects(1).Chart.SeriesCollection(1)
    Ser.HasDataLabels = true
    Ser.DataLabels.Position = xlLabelPositionBestFit
    Ser.HasLeaderLines = true
    Ser.LeaderLines.Border.ColorIndex = 5
}
```

```javascript
/*此示例显示图表工作表 Chart1 上的第二个数据系列是否有引导线。*/
function test() {
    let series2 = Application.Charts.Item("Chart1").ChartObjects(1).Chart.SeriesCollection(2)
    console.log(series2.HasLeaderLines)
}
```
