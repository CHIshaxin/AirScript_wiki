---
title: "BubbleSizes 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Series/%E5%B1%9E%E6%80%A7/BubbleSizes%20%E5%B1%9E%E6%80%A7.html
---
# BubbleSizes (属性)

返回或设置一个字符串，该字符串引用包含气泡图 x 值、y 值和大小数据的工作表单元格。在返回单元格引用时，它将返回以 A1 样式注释描述单元格的字符串。要为气泡图设置大小数据，必须使用 R1 样式注释。仅适用于气泡图。**Variant** 类型，可读写。

## 示例

```javascript
/*本示例显示图表系列中包含的气泡图 x 值、y 值和大小数据。*/
function test() {
    console.log(Worksheets.Item(1).ChartObjects(1).Chart.SeriesCollection(1).BubbleSizes.toString())
}
```

```javascript
/*此示例说明如何使用 R1C1 样式表示法设置BubbleSizes属性。*/
function test() {
    Application.Worksheets.Item(1).ChartObjects(1).Chart.SeriesCollection(1).BubbleSizes = "=Sheet1!r1c5:r5c5"
}
```
