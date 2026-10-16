---
title: "Points 方法"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Series/%E6%96%B9%E6%B3%95/Points%20%E6%96%B9%E6%B3%95.html
---
# Points (方法)

返回一个对象，该对象表示数据系列中单个数据点（ **[Point](./../../Point/Point 对象.html)** 对象）或所有数据点的集合（ **[Points](./../../Points/Points 对象.html)** 集合）。只读

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| Index | any | 可选 | 数据点的名称或编号。 |

## 返回值

Object

## 示例

```javascript
/*本示例对 Chart1 中第一个数据系列的第一个数据点应用数据标签。*/
function test() {
    Application.Charts.Item("Chart1").ChartObjects(1).Chart.SeriesCollection().Item(1).Points(1).ApplyDataLabels()
}
```

```javascript
/*此示例显示第一张工作表上的第一个图表的第二个数据系列的第二个数据点宽度。*/
function test() {
    let series2 = Application.Worksheets.Item(1).ChartObjects(1).Chart.SeriesCollection(2)
    console.log(series2.Points(2).Width)
}
```
