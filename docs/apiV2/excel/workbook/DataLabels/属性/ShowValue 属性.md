---
title: "ShowValue 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/DataLabels/%E5%B1%9E%E6%80%A7/ShowValue%20%E5%B1%9E%E6%80%A7.html
---
# ShowValue (属性)

返回或设置一个 **Boolean** 值，该值对应所指定图表的数据标签值显示方式。如果为 **True**，则显示数据标签值。如果为 **False**，则隐藏数据标签值。可读写。

## 说明

以程序方式访问数据标签之前，必须先激活指定的图表，否则将出现运行错误。

## 示例

```javascript
/*此示例显示活动工作表中第一个图表上第一个数据系列的数据标签的数据标签值。*/
function test() {
    ActiveSheet.ChartObjects(1).Activate()
    ActiveChart.SeriesCollection(1).DataLabels().ShowValue = true
}
```

```javascript
/*此示例将图表工作表 Chart1 中图表上的第二个数据系列的数据标签的值设置为不可见。*/
function test() {
    let datalabels = Application.Charts.Item("Chart1").ChartObjects(1).Chart.SeriesCollection(2).DataLabels()
    datalabels.ShowValue = false
}
```
