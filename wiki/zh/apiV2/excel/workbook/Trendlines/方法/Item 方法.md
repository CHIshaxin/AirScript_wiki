---
title: "Item (方法)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Trendlines/%E6%96%B9%E6%B3%95/Item%20%E6%96%B9%E6%B3%95.html
updated: 1789700916000
hash: DzC82qIf
---

# Item (方法)

从集合中返回一个对象。

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| Index | any | 可选 | 对象的索引号。 |

## 返回值

包含在集合中的一个 Trendline 对象。

## 示例

```javascript
/*本示例对 Chart1 中的趋势线向前和向后延伸的单位数进行设置。本示例应在包含具有趋势线的单个序列的 2D 柱形图上运行。*/
function test() {
    let trendline = Application.Charts.Item("Chart1").ChartObjects(1).Chart.SeriesCollection(1).Trendlines().Item(1)
    trendline.Forward = 5
    trendline.Backward = 0.5
}
```

```javascript
/*此示例显示第一张工作表上的第一个图表的第二个数据系列索引为2的趋势线名称。*/
function test() {
    let series = Application.Worksheets.Item(1).ChartObjects(1).Chart.SeriesCollection(2)
    console.log(series.Trendlines().Item(2).Name)
}
```
