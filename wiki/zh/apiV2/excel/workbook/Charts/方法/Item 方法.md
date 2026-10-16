---
title: "Item (方法)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Charts/%E6%96%B9%E6%B3%95/Item%20%E6%96%B9%E6%B3%95.html
updated: 1789700916000
hash: D79BDqGX
---

# Item (方法)

从集合中返回一个对象。

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
|---|---|---|---|
| Index | any | 必填 | 对象的名称或索引号。 |

## 示例

```javascript
/*此示例对图表工作表 Chart1 中的趋势线向前和向后延伸的单位数进行设置。此示例应在包含单个带趋势线系列的二维柱形图上运行。*/
function test() {
    let chartLine = Application.Charts.Item("Chart1").SeriesCollection(1).Trendlines(1)
    chartLine.Forward2 = 5
    chartLine.Backward2 = 0.5
}
```

```javascript
/*本示例显示图表工作表 Chart1 的图例，然后将图例的字体颜色设为蓝色。*/
function test() {
    let chart = Application.Charts.Item("Chart1")
    chart.HasLegend = true
    chart.Legend.Font.ColorIndex = 5
}
```
