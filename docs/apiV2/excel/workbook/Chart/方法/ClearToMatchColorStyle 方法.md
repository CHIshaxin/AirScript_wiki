---
title: "ClearToMatchColorStyle 方法"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Chart/%E6%96%B9%E6%B3%95/ClearToMatchColorStyle%20%E6%96%B9%E6%B3%95.html
---
# ClearToMatchColorStyle (方法)

清除指定图表上不遵循应用于图表的颜色样式的所有颜色。

## 说明

使用此方法将所有图表颜色格式重置为应用于图表的颜色样式。

## 示例

```javascript
/*本示例重置工作表 Sheet1 第一张内嵌图表的所有颜色元素格式。*/
function test() {
    let chart = Application.Worksheets.Item("Sheet1").ChartObjects(1).Chart
    chart.ClearToMatchColorStyle()
}
```

```javascript
/*本示例重置图表工作表 Chart1 所有颜色元素格式。*/
function test() {
    let chart = Application.Charts.Item("Chart1").ChartObjects(1).Chart
    chart.ClearToMatchColorStyle()
}
```
