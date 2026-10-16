---
title: "SeriesCollection (对象)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/SeriesCollection/SeriesCollection%20%E5%AF%B9%E8%B1%A1.html
updated: 1789700916000
hash: bJ0aYbQ9
---

# SeriesCollection (对象)

指定的图表或图表组中所有 **[Series](./../Series/Series 对象.html)** 对象的集合。

## 说明

使用 **SeriesCollection**方法可返回 **[SeriesCollection](./../SeriesCollection/SeriesCollection 对象.html)** 集合。

## 示例

```javascript
/*此示例显示第一张工作表中第一个图表的系列数量。*/
function test() {
    console.log(Application.Worksheets.Item(1).ChartObjects(1).Chart.SeriesCollection().Count)
}
```

```javascript
/*此示例将单元格 A1:A19 中的数据添加到名为 Chart1 的图表工作表上的新系列。*/
function test() {
    Charts.Item("Chart1").ChartObjects(1).Chart.SeriesCollection().Add(Worksheets.Item("Sheet1").Range("A1:A19"))
}
```

```javascript
/*此示例设置工作表 Sheet1 上嵌入的第一个图表中第一个数据系列的内部颜色。*/
function test() {
    Worksheets("Sheet1").ChartObjects(1).Chart.SeriesCollection(1).Interior.Color = RGB(255, 0, 0)
}
```
