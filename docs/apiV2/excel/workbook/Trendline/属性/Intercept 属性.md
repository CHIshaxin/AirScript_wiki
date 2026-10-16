---
title: "Intercept 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Trendline/%E5%B1%9E%E6%80%A7/Intercept%20%E5%B1%9E%E6%80%A7.html
---
# Intercept (属性)

返回或设置趋势线与数值轴的交叉点。可读/写 **Double** 类型。

## 说明

设置该属性会将 **InterceptIsAuto** 属性设置为 **False**。

## 示例

```javascript
/*本示例设置 Chart1 的第一条趋势线与数值轴相交于 5。本示例应在包含单个带趋势线系列的二维柱形图上运行。*/
function test() {
    Application.Charts.Item("Chart1").ChartObjects(1).Chart.SeriesCollection(1).Trendlines(1).Intercept = 5
}
```

```javascript
/*此示例显示第一张工作表上的第一个图表的第二个数据系列索引为2的趋势线与数值轴的交叉点。*/
function test() {
    let series = Application.Worksheets.Item(1).ChartObjects(1).Chart.SeriesCollection(2)
    console.log(series.Trendlines(2).Intercept)
}
```
