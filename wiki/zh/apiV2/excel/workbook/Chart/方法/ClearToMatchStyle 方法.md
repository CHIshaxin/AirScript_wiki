---
title: "ClearToMatchStyle (方法)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Chart/%E6%96%B9%E6%B3%95/ClearToMatchStyle%20%E6%96%B9%E6%B3%95.html
updated: 1789700916000
hash: BiHQSbYo
---

# ClearToMatchStyle (方法)

清除图表元素格式以改为自动格式。

## 说明

使用此方法将图表元素格式重置为自动格式。如果在图表上使用此方法，则将重置所有格式（包括替代项）。

## 示例

```javascript
/*本示例清除第一张工作表第一张内嵌图表的元素格式以改为自动格式。*/
function test() {
    let chart = Application.Worksheets.Item(1).ChartObjects(1).Chart
    chart.ClearToMatchStyle()
}
```

```javascript
/*本示例清除图表工作表 Chart1 元素格式以改为自动格式。*/
function test() {
    Application.Charts.Item("Chart1").ClearToMatchStyle()
}
```
