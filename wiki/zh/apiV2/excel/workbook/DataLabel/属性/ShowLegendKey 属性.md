---
title: "ShowLegendKey (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/DataLabel/%E5%B1%9E%E6%80%A7/ShowLegendKey%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: CJiPh9Th
---

# ShowLegendKey (属性)

如果数据标签图例项标示可见，则为 **True**。**Boolean** 类型，可读写。

## 示例

```javascript
/*此示例设置 Chart1 中系列一的第三个数据标签，以显示数值和图例标示。*/
function test() {
    let datalabel = Application.Charts.Item("Chart1").ChartObjects(1).Chart.SeriesCollection(1).DataLabels(3)
    datalabel.ShowLegendKey = true
    datalabel.ShowValue = true
}
```

```javascript
/*此示例判断工作表 Sheet1 中第一张图表的第二个系列的第三个数据标签的图例项标示是否可见，并通知用户。*/
function test() {
    let datalabel = Application.Sheets.Item("Sheet1").ChartObjects(1).Chart.SeriesCollection(2).DataLabels(3)
    if (datalabel.ShowLegendKey) {
        console.log("数据标签的图例项标示可见")
    } else {
        console.log("数据标签的图例项标示不可见")
    }
}
```
