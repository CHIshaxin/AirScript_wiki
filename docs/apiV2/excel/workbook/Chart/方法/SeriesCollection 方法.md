---
title: "SeriesCollection 方法"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Chart/%E6%96%B9%E6%B3%95/SeriesCollection%20%E6%96%B9%E6%B3%95.html
---
# SeriesCollection (方法)

返回一个对象，它代表图表或图表组中的一个系列（ **[Series](./../../Series/Series 对象.html)** 对象）或所有系列的集合（ **[SeriesCollection](./../../SeriesCollection/SeriesCollection 对象.html)** 集合）。

## 示例

```javascript
/*此示例打开图表工作表 Chart1 中第一个数据系列的数据标签。*/
function test() {
    let chart = Application.Charts.Item("Chart1").ChartObjects(1).Chart
    chart.SeriesCollection(1).HasDataLabels = true
}
```

```javascript
/*此示例设置图表工作表 Chart1 中第二个数据系列的内部颜色为红色。*/
function test() {
    let chart = Application.Charts.Item("Chart1").ChartObjects(1).Chart
    chart.SeriesCollection(2).Interior.Color = RGB(255, 0, 0)
}
```

param Index - 数据系列的名称或编号。
