---
title: "NewSeries (方法)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/SeriesCollection/%E6%96%B9%E6%B3%95/NewSeries%20%E6%96%B9%E6%B3%95.html
updated: 1789700916000
hash: BxbjLRJY
---

# NewSeries (方法)

创建新系列。返回代表该新系列的 **[Series](./../../Series/Series 对象.html)** 对象。

## 说明

本方法对数据透视图无效。

## 返回值

Series

## 示例

```javascript
/*本示例向图表工作表 Chart1 中添加新系列。*/
function test() {
    let ns = Application.Charts.Item("Chart1").ChartObjects(1).Chart.SeriesCollection().NewSeries()
}
```

```javascript
/*此示例向第一张工作表中第一个图表添加新系列。*/
function test() {
    Application.Worksheets.Item(1).ChartObjects(1).Chart.SeriesCollection().NewSeries()
}
```
