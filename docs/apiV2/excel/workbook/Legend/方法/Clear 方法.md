---
title: "Clear 方法"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Legend/%E6%96%B9%E6%B3%95/Clear%20%E6%96%B9%E6%B3%95.html
---
# Clear (方法)

清除整个对象。

## 返回值

Variant

## 示例

```javascript
/*本示例将工作表一上嵌入图表一的图例清除。*/
function test() {
    Worksheets.Item(1).ChartObjects(1).Chart.Legend.Clear()
}
```

```javascript
/*此示例清除图表工作表 Chart1 上的第一个图表中的图例。*/
function test() {
    Application.Charts.Item("Chart1").ChartObjects(1).Chart.Legend.Clear()
}
```
