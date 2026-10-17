---
title: "NameIsAuto (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Trendline/%E5%B1%9E%E6%80%A7/NameIsAuto%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: C4jYzsSj
---

# NameIsAuto (属性)

如果 ET 自动确定趋势线的名称，则该属性值为 **True**。**Boolean** 类型，可读写。

## 示例

```javascript
/*本示例设置 ET 自动确定 Chart1 的趋势线一的名称。本示例应在包含单个带趋势线系列的二维柱形图上运行。*/
function test() {
    Application.Charts.Item("Chart1").ChartObjects(1).Chart.SeriesCollection(1).Trendlines(1).NameIsAuto = true
}
```

```javascript
/*此示例显示ET是否能自动确定第一张工作表上的第一个图表的第二个数据系列索引为2的趋势线名称。*/
function test() {
    let series = Application.Worksheets.Item(1).ChartObjects(1).Chart.SeriesCollection(2)
    console.log(series.Trendlines(2).NameIsAuto)
}
```
