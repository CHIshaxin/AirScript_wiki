---
title: "ShowBubbleSize 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/DataLabel/%E5%B1%9E%E6%80%A7/ShowBubbleSize%20%E5%B1%9E%E6%80%A7.html
---
# ShowBubbleSize (属性)

如果为 **True**，则在图表中显示数据标签的气泡的大小。如果为 **False**，则隐藏气泡的大小。**Boolean**类型，可读写。

## 说明

以程序方式访问数据标签之前，必须先激活图表，否则将出现运行错误。

## 示例

```javascript
/*此示例显示活动工作表第一个图表上第一个系列的第一个数据标签的气泡大小。此示例假定活动工作表上存在图表。*/
function test() {
    Application.ActiveSheet.ChartObjects(1).Activate()
    ActiveChart.SeriesCollection(1).Points(1).DataLabel.ShowBubbleSize = true
}
```

```javascript
/*此示例显示图表工作表 Chart1 上第一个系列的第二个数据标签是否展示气泡大小。*/
function test() {
    let datalabel = Application.Charts.Item("Chart1").ChartObjects(1).Chart.SeriesCollection(1).Points(2).DataLabel
    console.log(datalabel.ShowBubbleSize)
}
```
