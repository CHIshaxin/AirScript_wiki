---
title: "ClearFormats 方法"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/LegendKey/%E6%96%B9%E6%B3%95/ClearFormats%20%E6%96%B9%E6%B3%95.html
---
# ClearFormats (方法)

清除对象的格式设置。

## 返回值

Variant

## 示例

```javascript
/*下例清除 Sheet1 上第一个图表的第三个图例项的图例标示格式。*/
function test() {
    let legend = Application.Worksheets.Item("Sheet1").ChartObjects(1).Chart.Legend
    legend.LegendEntries(3).LegendKey.ClearFormats()
}
```

```javascript
/*本示例清除图表工作表 Chart1 上第二个图例项的图例标示的格式。*/
function test() {
    let legend = Application.Charts.Item("Chart1").ChartObjects(1).Chart.Legend
    legend.LegendEntries(2).LegendKey.ClearFormats()
}
```
