---
title: "Shadow 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/DataLabels/%E5%B1%9E%E6%80%A7/Shadow%20%E5%B1%9E%E6%80%A7.html
---
# Shadow (属性)

返回或设置一个 **Boolean** 值，它确定对象是否有阴影。

## 示例

```javascript
/*此示例将图表工作表 Chart1 中图表的第一个数据序列的数据标签设置为有阴影。*/
function test() {
    let datalabels = Application.Charts.Item("Chart1").ChartObjects(1).Chart.SeriesCollection(1).DataLabels()
    datalabels.Shadow = true
}
```

```javascript
/*此示例显示工作表 Sheet1 中第一个图表上的第三个系列的数据标签是否有阴影。*/
function test() {
    let datalabel = Application.Worksheets.Item("Sheet1").ChartObjects(1).Chart.SeriesCollection(3).DataLabels()
    console.log(datalabel.Shadow)
}
```
