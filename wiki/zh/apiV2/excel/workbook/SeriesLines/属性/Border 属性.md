---
title: "Border (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/SeriesLines/%E5%B1%9E%E6%80%A7/Border%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: DWTg_kbo
---

# Border (属性)

返回一个 **[Border](./../../Border/Border 对象.html)** 对象，它代表对象的边框。

## 示例

```javascript
/*本示例打开图表工作表 Chart1 上第一个图表组的系列线，然后设置其线条颜色为蓝色。*/
function test() {
    let chartgroup = Application.Charts.Item("Chart1").ChartObjects(1).Chart.ChartGroups(1)
    chartgroup.HasSeriesLines = true
    let border = chartgroup.SeriesLines.Border
    border.Color = RGB(0, 0, 255)
}
```

```javascript
/*本示例打开工作表 Sheet1 第一个内嵌图表上第一个图表组的系列线，并显示其是否为细线。该示例应在具有两个或更多系列的 2D 堆积柱形图上运行。*/
function test() {
    let chartgroup = Application.Sheets.Item("Sheet1").ChartObjects(1).Chart.ChartGroups(1)
    chartgroup.HasSeriesLines = true
    console.log(chartgroup.SeriesLines.Border.Weight == xlHairline)
}
```
