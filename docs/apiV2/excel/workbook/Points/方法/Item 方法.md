---
title: "Item 方法"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Points/%E6%96%B9%E6%B3%95/Item%20%E6%96%B9%E6%B3%95.html
---
# Item (方法)

从集合中返回一个对象。

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| Index | int4 | 必填 | 对象的索引号。 |

## 返回值

包含在集合中的一个 **[Point](./../../Point/Point 对象.html)** 对象。

## 示例

```javascript
/*本示例设置第一个工作表上嵌入图表第一个系列中的第三个点的标记样式。指定的系列必须是 2D 线、散点或雷达系列。*/
function test() {
    Worksheets.Item(1).ChartObjects(1).Chart.SeriesCollection(1).Points().Item(3).MarkerStyle = xlDiamond
}
```

```javascript
/*本示例显示 Chart1 中第三个数据系列的第4个数据点的宽度。*/
function test() {
    let series = Application.Charts.Item("Chart1").ChartObjects(1).Chart.SeriesCollection(3)
    console.log(series.Points().Item(4).Width)
}
```
