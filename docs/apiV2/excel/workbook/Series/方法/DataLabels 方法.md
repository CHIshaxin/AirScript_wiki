---
title: "DataLabels 方法"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Series/%E6%96%B9%E6%B3%95/DataLabels%20%E6%96%B9%E6%B3%95.html
---
# DataLabels (方法)

返回代表数据系列中单个数据标签（ **[DataLabel](./../../DataLabel/DataLabel 对象.html)** 对象）或所有数据标签的集合（ **[DataLabels](./../../DataLabels/DataLabels 对象.html)** 集合）的对象。

## 说明

如果指定数据标签的“值”选项处于打开状态，则返回的集合中每个数据点最多包含一个数据标签。可分别设置系列中数据点的数据标签开关选项。

如果指定系列处于面积图中，并且数据标签的“系列名称”选项处于打开状态，则返回的集合中仅包含单个数据标签，即面积系列的标签。

## 示例

```javascript
/*本示例对 Chart1 的第一个数据系列的数据标签进行设置，显示其关键字段的值。本示例假定在运行时这些值是可见的。*/
function test() {
    let series = Application.Charts.Item("Chart1").ChartObjects(1).Chart.SeriesCollection().Item(1)
    series.HasDataLabels = true
    series.DataLabels().ShowLegendKey = true
}
```

```javascript
/*此示例显示活动工作表上的第一个图表的第二个系列的第一个标签宽度。*/
function test() {
    console.log(ActiveSheet.ChartObjects(1).Chart.SeriesCollection(2).DataLabels(1).Width)
}
```
