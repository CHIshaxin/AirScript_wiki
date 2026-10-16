---
title: "Has3DEffect (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Point/%E5%B1%9E%E6%80%A7/Has3DEffect%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: BWNLcT9h
---

# Has3DEffect (属性)

如果数据点具有三维外观，则为 **True**。可读/写 **Boolean** 类型。

## 说明

该属性仅适用于气泡图。

## 示例

```javascript
/*此示例设置第一张工作表上第一个图表的第三个数据系列中第3个数据点具有三维外观。*/
function test() {
    let series = Application.Worksheets.Item(1).ChartObjects(1).Chart.SeriesCollection(3)
    series.Points(3).Has3DEffect = true
}
```

```javascript
/*本示例显示 Chart1 中第二个数据系列的第2个数据点是否具有三维外观。*/
function test() {
    let series = Application.Charts.Item("Chart1").ChartObjects(1).Chart.SeriesCollection(2)
    console.log(series.Points(2).Has3DEffect)
}
```
