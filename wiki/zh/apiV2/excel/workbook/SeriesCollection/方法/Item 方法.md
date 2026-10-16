---
title: "Item (方法)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/SeriesCollection/%E6%96%B9%E6%B3%95/Item%20%E6%96%B9%E6%B3%95.html
updated: 1789700916000
hash: GdY57stJ
---

# Item (方法)

从集合中返回一个对象。

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| Index | any | 必填 | 对象的名称或索引号。 |

## 返回值

包含在集合中的一个 Series 对象。

## 示例

```javascript
/*本示例对 Chart1 中的趋势线向前和向后延伸的单位数进行设置。本示例应在包含单个带趋势线系列的二维柱形图上运行。*/
function test() {
    let trendlines = Application.Charts.Item("Chart1").ChartObjects(1).Chart.SeriesCollection().Item(1).Trendlines().Item(1)
    trendlines.Forward = 5
    trendlines.Backward = .5
}
```

```javascript
/*此示例将第一张工作表上的第一个图表的第二个数据系列名称设置为“系列二”。*/
function test() {
    let series2 = Application.Worksheets.Item(1).ChartObjects(1).Chart.SeriesCollection().Item(2)
    series2.Name = "系列二"
}
```
