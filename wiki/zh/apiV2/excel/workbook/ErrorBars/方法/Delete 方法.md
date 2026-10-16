---
title: "Delete (方法)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/ErrorBars/%E6%96%B9%E6%B3%95/Delete%20%E6%96%B9%E6%B3%95.html
updated: 1789700916000
hash: diOI7OYm
---

# Delete (方法)

删除对象。

## 返回值

Variant

## 示例

```javascript
/*此示例删除工作表 Sheet1 中第一张图表的第一个系列的误差线。*/
function test() {
    let errorbars = Application.Worksheets.Item("Sheet1").ChartObjects(1).Chart.SeriesCollection(1).ErrorBars
    errorbars.Delete()
}
```

```javascript
/*此示例删除图表工作表 Chart1 中图表的第三个系列的误差线。*/
function test() {
    let errorbars = Application.Charts.Item("Chart1").ChartObjects(1).Chart.SeriesCollection(3).ErrorBars
    errorbars.Delete()
}
```
