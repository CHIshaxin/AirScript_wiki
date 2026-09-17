---
title: "ShowSeriesName 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/DataLabel/%E5%B1%9E%E6%80%A7/ShowSeriesName%20%E5%B1%9E%E6%80%A7.html
---
# ShowSeriesName (属性)

返回或设置一个 **Boolean** 值，它指明图表上数据标签的系列名称显示方式。如果为 **True**，则显示系列名称。如果为 **False**，则隐藏系列名称。可读写。

## 说明

以程序方式访问数据标签之前，必须先激活图表，否则将出现运行错误。

## 示例

```javascript
/*此示例显示活动工作表第一个图表上第一个数据系列的第二个数据标签的系列名称。此示例假定活动工作表上存在图表。*/
function test() {
    Application.ActiveSheet.ChartObjects(1).Activate()
    ActiveChart.SeriesCollection(1).DataLabels(2).ShowSeriesName = true
}
```

```javascript
/*此示例显示图表工作表 Chart1 中图表的第一个系列的第二个数据标签是否开启系列名称。*/
function test() {
    let datalabel = Application.Charts.Item("Chart1").ChartObjects(1).Chart.SeriesCollection(1).DataLabels(2)
    console.log(datalabel.ShowSeriesName)
}
```
