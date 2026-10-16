---
title: "ShowValue (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/DataLabel/%E5%B1%9E%E6%80%A7/ShowValue%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: DDOb8LYR
---

# ShowValue (属性)

返回或设置一个 **Boolean** 值，该值对应所指定图表的数据标签值显示方式。如果为 **True**，则显示数据标签值。如果为 **False**，则隐藏数据标签值。可读写。

## 说明

以程序方式访问数据标签之前，必须先激活指定的图表，否则将出现运行错误。

## 示例

```javascript
/*此示例显示活动工作表第一个图表上第一个数据系列的第二个数据标签的值。此示例假定活动工作表上存在图表。*/
function test() {
    Application.ActiveSheet.ChartObjects(1).Activate()
    ActiveChart.SeriesCollection(1).DataLabels(2).ShowValue = true
}
```

```javascript
/*此示例显示图表工作表 Chart1 中图表的第二个系列的第二个数据标签是否展示数据标签值。*/
function test() {
    let datalabel = Application.Charts.Item("Chart1").ChartObjects(1).Chart.SeriesCollection(2).DataLabels(2)
    console.log(datalabel.ShowValue)
}
```
