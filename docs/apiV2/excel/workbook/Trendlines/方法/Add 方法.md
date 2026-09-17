---
title: "Add 方法"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Trendlines/%E6%96%B9%E6%B3%95/Add%20%E6%96%B9%E6%B3%95.html
---
# Add (方法)

创建新的趋势线。

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| Order | any | 可选 | 如果 Type 为 xlPolynomial。趋势线顺序号。必须为 2 到 6 之间的整数（包括 2 和 6）。 |
| Period | any | 可选 | 如果 Type 为 xlMovingAvg。趋势线周期。必须为大于 1，而小于正添加趋势线的数据系列中数据点个数的整数。 |
| Forward | any | 可选 | 趋势线向前延伸的周期数目（或散点图中的单位个数）。 |
| Backward | any | 可选 | 趋势线向后延伸的周期数目（或散点图中的单位个数）。 |
| Intercept | any | 可选 | 趋势线的截距。如果省略此参数，则由回归分析自动设置截距。 |
| DisplayEquation | any | 可选 | 如果为 True，则在图表中显示趋势线的公式（与 R 平方值显示在同一数据标签中）。默认值为 False。 |
| DisplayRSquared | any | 可选 | 如果为 True，则在图表中显示趋势线的 R 平方值（与公式显示在同一数据标签中）。默认值为 False。 |
| Name | any | 可选 | 作为文本的趋势线名称。如果省略此参数，则由 ET 生成名称。 |

## 返回值

一个代表新趋势线的 Trendline 对象。

## 示例

```javascript
/*本示例在 Chart1 上创建新的线性趋势线。*/
function test() {
    Application.ActiveWorkbook.Charts.Item("Chart1").ChartObjects(1).Chart.SeriesCollection(1).Trendlines().Add()
}
```

```javascript
/*此示例在第一张工作表上的第一个图表的第二个数据系列创建新的趋势线。*/
function test() {
    let series = Application.Worksheets.Item(1).ChartObjects(1).Chart.SeriesCollection(2)
    series.Trendlines().Add(xlMovingAvg, 3, 4, 2, true, true, true)
}
```
