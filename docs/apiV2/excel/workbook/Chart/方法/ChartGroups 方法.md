---
title: "ChartGroups 方法"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Chart/%E6%96%B9%E6%B3%95/ChartGroups%20%E6%96%B9%E6%B3%95.html
---
# ChartGroups (方法)

返回一个对象，该对象表示图表中单个图表组（ **[ChartGroup](./../../ChartGroup/ChartGroup 对象.html)** 对象）或所有图表组的集合（ **[ChartGroups](./../../ChartGroups/ChartGroups 对象.html)** 对象）。返回的集合中包括每种类型的图表组。

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| Index | any | 可选 | 图表组编号。 |

## 示例

```javascript
/*本示例显示图表工作表 Chart1 中第一个图表组的涨跌柱线，并对其颜色进行设置。该示例应在包含两个序列的 2D 折线图上运行，这些序列与一个或多个数据点相交。*/
function test() {
    let chartgroup = Application.Charts.Item("Chart1").ChartObjects(1).Chart.ChartGroups(1)
    chartgroup.HasUpDownBars  = true
    chartgroup.DownBars.Interior.ColorIndex = 3
    chartgroup.UpBars.Interior.ColorIndex = 5
}
```

```javascript
/*此示例为工作表 Sheet1 中第一个内嵌图表的第一个图表组添加垂直线。*/
function test() {
    let chart = Application.Worksheets.Item("Sheet1").ChartObjects(1).Chart
    chart.ChartGroups(1).HasDropLines = true
}
```
