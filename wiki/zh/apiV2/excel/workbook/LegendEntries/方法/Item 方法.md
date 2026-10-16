---
title: "Item (方法)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/LegendEntries/%E6%96%B9%E6%B3%95/Item%20%E6%96%B9%E6%B3%95.html
updated: 1789700916000
hash: BwEeW-Tl
---

# Item (方法)

从集合中返回一个对象。

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| Index | any | 必填 | 对象的索引号。 |

## 返回值

包含在集合中的一个 LegendEntry 对象。

## 示例

```javascript
/*此示例对 Sheet1 中第一张嵌入式图表的图例（通常是第一个数据系列的图例）顶部的图例项的文本字体进行更改。*/
function test() {
    Application.Worksheets.Item("Sheet1").ChartObjects(1).Chart.Legend.LegendEntries().Item(1).Font.Italic = true
}
```

```javascript
/*本示例删除 Chart1 上第一个图例项。*/
function test() {
    Application.Charts.Item("Chart1").ChartObjects(1).Chart.Legend.LegendEntries().Item(1).Delete()
}
```
