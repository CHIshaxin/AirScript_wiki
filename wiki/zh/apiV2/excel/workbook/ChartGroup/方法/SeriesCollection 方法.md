---
title: "SeriesCollection (方法)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/ChartGroup/%E6%96%B9%E6%B3%95/SeriesCollection%20%E6%96%B9%E6%B3%95.html
updated: 1789700916000
hash: DR9ER29F
---

# SeriesCollection (方法)

返回一个对象，它代表图表或图表组中的一个系列（ **[Series](./../../Series/Series 对象.html)** 对象）或所有系列的集合（ **[SeriesCollection](./../../SeriesCollection/SeriesCollection 对象.html)** 集合）。

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
|---|---|---|---|
| Index | any | 可选 | 数据系列的名称或编号。 |

## 示例

```javascript
/*此示例打开图表工作表 Chart1 中第一个数据系列的数据标签。*/
function test() {
    let chart = Application.Charts.Item("Chart1").ChartObjects(1).Chart
    chart.SeriesCollection(1).HasDataLabels = true
}
```

```javascript
/*此示例设置工作表 Sheet1 第一个内嵌图表中第一个图表组的第二个数据系列的内部颜色为红色。*/
function test() {
    let chartgroup = Application.Sheets.Item("Sheet1").ChartObjects(1).Chart.ChartGroups(1)
    chartgroup.SeriesCollection(2).Interior.Color = RGB(255, 0, 0)
}
```
