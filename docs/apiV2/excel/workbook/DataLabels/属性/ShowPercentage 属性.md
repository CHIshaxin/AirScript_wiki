---
title: "ShowPercentage 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/DataLabels/%E5%B1%9E%E6%80%A7/ShowPercentage%20%E5%B1%9E%E6%80%A7.html
---
# ShowPercentage (属性)

如果为 **True**，则在图表中显示数据标签的百分比数值。如果为 **False**，则隐藏该百分比数值。**Boolean**类型，可读写。

## 说明

以程序方式访问数据标签之前，必须先激活图表，否则将出现运行错误。

## 示例

```javascript
/*此示例显示活动工作表中第一个图表上第一个数据系列的数据标签的百分比数值。此示例假定活动工作表上存在图表且图表为饼图。*/
function test() {
    Application.ActiveSheet.ChartObjects(1).Activate()
    ActiveChart.SeriesCollection(1).DataLabels().ShowPercentage = true
}
```

```javascript
/*此示例判断图表工作表 Chart1 中图表的第一个数据系列的数据标签的百分比数值是否可见，并通知用户。*/
function test() {
    let datalabels = Application.Charts.Item("Chart1").ChartObjects(1).Chart.SeriesCollection(1).DataLabels()
    if (datalabels.ShowPercentage) {
        console.log("数据标签的百分比数值可见")
    } else {
        console.log("数据标签的百分比数值不可见")
    }
}
```
