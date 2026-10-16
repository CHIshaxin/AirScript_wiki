---
title: "Trendlines (方法)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Series/%E6%96%B9%E6%B3%95/Trendlines%20%E6%96%B9%E6%B3%95.html
updated: 1789700916000
hash: BBB40zRU
---

# Trendlines (方法)

返回一个对象，该对象表示单个趋势线（ **[Trendline](./../../Trendline/Trendline 对象.html)** 对象）或所有趋势线的集合（ **[Trendlines](./../../Trendlines/Trendlines 对象.html)** 集合）。

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| Index | any | 可选 | 趋势线的名称或编号。 |

## 返回值

Object

## 示例

```javascript
/*本示例向 Chart1 中第一个数据系列添加线性趋势线。*/
function test() {
    Application.Charts.Item("Chart1").ChartObjects(1).Chart.SeriesCollection().Item(1).Trendlines().Add(xlLinear)
}
```

```javascript
/*此示例显示第一张工作表上的第一个图表的第二个数据系列索引为1的趋势线名字。*/
function test() {
    let series2 = Application.Worksheets.Item(1).ChartObjects(1).Chart.SeriesCollection(2)
    console.log(series2.Trendlines(1).Name)
}
```
