---
title: "Delete 方法"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/SeriesLines/%E6%96%B9%E6%B3%95/Delete%20%E6%96%B9%E6%B3%95.html
---
# Delete (方法)

删除对象。

## 返回值

Variant

## 示例

```javascript
/*本示例将图表工作表 Chart1 上第一个图表组的系列线删除。*/
function test() {
    let chartgroup = Application.Charts.Item("Chart1").ChartObjects(1).Chart.ChartGroups(1)
    chartgroup.SeriesLines.Delete()
}
```

```javascript
/*本示例将工作表 Sheet1 第一个内嵌图表上第一个图表组的系列线删除，并显示是否删除成功。*/
function test() {
    let chartgroup = Application.Sheets.Item("Sheet1").ChartObjects(1).Chart.ChartGroups(1)
    console.log(chartgroup.SeriesLines.Delete())
}
```
