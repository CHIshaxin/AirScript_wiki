---
title: "Name 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/ErrorBars/%E5%B1%9E%E6%80%A7/Name%20%E5%B1%9E%E6%80%A7.html
---
# Name (属性)

返回一个 **String** 值，它代表对象的名称。

## 示例

```javascript
/*本示例显示图表工作表 Chart1 中图表的第二个数据系列的误差线的名称。*/
function test() {
    let errorbars = Application.Charts.Item("Chart1").ChartObjects(1).Chart.SeriesCollection(2).ErrorBars
    console.log(errorbars.Name)
}
```

```javascript
/*本示例显示工作表 Sheet1 中第一个图表的第三个系列的误差线的名称。*/
function test() {
    let errorbars = Application.Sheets.Item("Sheet1").ChartObjects(1).Chart.SeriesCollection(3).ErrorBars
    console.log(`误差线的名称：${errorbars.Name}`)
}
```
