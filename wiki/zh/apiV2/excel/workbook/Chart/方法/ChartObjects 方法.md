---
title: "ChartObjects (方法)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Chart/%E6%96%B9%E6%B3%95/ChartObjects%20%E6%96%B9%E6%B3%95.html
updated: 1789700916000
hash: 6QtI8UZv
---

# ChartObjects (方法)

返回一个对象，它代表工作表上的一个嵌入式图表（ **[ChartObject](./../../ChartObject/ChartObject 对象.html)** 对象）或所有嵌入式图表的集合（ **[ChartObjects](./../../ChartObjects/ChartObjects 对象.html)** 对象）。

## 说明

该方法不等同于 **[Charts](./../../Charts/Charts 对象.html)** 属性，它返回嵌入式图表，而 **[Charts](./../../Charts/Charts 对象.html)** 属性返回图表工作表。使用 **Chart** 属性可返回嵌入式图表的 **[Chart](./../../Chart/Chart 对象.html)** 对象。

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
|---|---|---|---|
| Index | any | 可选 |  |

## 示例

```javascript
/*本示例向图表工作表 Chart1 上的第一个嵌入式图表中添加标题。*/
function test() {
    let chart = Application.Charts.Item("Chart1").ChartObjects(1).Chart
    chart.HasTitle = true
    chart.ChartTitle.Text = "1995 Rainfall Totals by Month"
}
```

```javascript
/*本示例在工作表 Sheet1 上的第一个嵌入式图表中新建一个系列，新系列的数据源为 Sheet1 的 B1:B10 区域。*/
function test() {
    Application.Worksheets.Item("Sheet1").ChartObjects(1).Activate()
    Application.ActiveChart.SeriesCollection().Add(Worksheets.Item("Sheet1").Range("B1:B10"))
}
```

```javascript
/*本示例清除工作表 Sheet1 上第一个嵌入式图表的格式设置。*/
function test() {
    Application.Worksheets.Item("Sheet1").ChartObjects(1).Chart.ChartArea.ClearFormats()
}
```
