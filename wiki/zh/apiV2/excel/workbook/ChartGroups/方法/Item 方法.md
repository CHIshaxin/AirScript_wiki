---
title: "Item (方法)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/ChartGroups/%E6%96%B9%E6%B3%95/Item%20%E6%96%B9%E6%B3%95.html
updated: 1789700916000
hash: CCNm1rl_
---

# Item (方法)

从集合中返回一个对象。

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
|---|---|---|---|
| Index | any | 必填 | 对象的索引号。 |

## 示例

```javascript
/*此示例为图表工作表 Chart1 中的第一个图表组添加垂直线。*/
function test() {
    let chartgroups = Application.Charts.Item("Chart1").ChartObjects(1).Chart.ChartGroups()
    chartgroups.Item(1).HasDropLines = true
}
```

```javascript
/*本示例将工作表 Sheet1 第一个内嵌图表中第一个图表组中的列簇间距设置为 250 。*/
function test() {
    let chartgroups = Application.Worksheets.Item("Sheet1").ChartObjects(1).Chart.ChartGroups()
    chartgroups.Item(1).GapWidth = 250
}
```
