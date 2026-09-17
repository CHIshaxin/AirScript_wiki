---
title: "Delete 方法"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/LegendEntry/%E6%96%B9%E6%B3%95/Delete%20%E6%96%B9%E6%B3%95.html
---
# Delete (方法)

删除对象。

## 返回值

Variant

## 示例

```javascript
/*本示例删除 Chart1 上第一个图例项。*/
function test() {
    Application.Charts.Item("Chart1").ChartObjects(1).Chart.Legend.LegendEntries().Item(1).Delete()
}
```

```javascript
/*此示例删除 Sheet1 中第二个图例项。*/
function test() {
    Application.Worksheets.Item("Sheet1").ChartObjects(1).Chart.Legend.LegendEntries().Item(2).Delete()
}
```
