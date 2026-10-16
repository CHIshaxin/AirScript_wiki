---
title: "IsFiltered (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/ChartCategory/%E5%B1%9E%E6%80%A7/IsFiltered%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: CigNDiut
---

# IsFiltered (属性)

当用户筛选出序列时，返回 True 。 只读 Boolean。

## 说明

当某个系列从其父 SeriesCollection 中移出时，该系列仍保留在其父级 FullSeriesCollection 中。 当用户重新筛选序列时，该序列将插入到 SeriesCollection 中的原始位置。

## 示例

```javascript
/*此示例显示图表工作表 Chart1 第一个图表组的第一个类别是否筛选出序列。*/
function test() {
    let chart = Application.Charts.Item("Chart1").ChartObjects(1).Chart.ChartGroups(1)
    console.log(chart.FullCategoryCollection(1).IsFiltered)
}
```

```javascript
/*此示例显示工作表 Sheet1 第一个内嵌图表第一个图表组第二个类别是否筛选出序列。*/
function test() {
    let chart = Application.Worksheets.Item("Sheet1").ChartObjects(1).Chart.ChartGroups(1)
    console.log(chart.CategoryCollection(2).IsFiltered)
}
```
