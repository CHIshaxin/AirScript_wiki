---
title: "NumberFormatLocal 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/DataLabels/%E5%B1%9E%E6%80%A7/NumberFormatLocal%20%E5%B1%9E%E6%80%A7.html
---
# NumberFormatLocal (属性)

以采用用户语言字符串的形式返回或设置一个 **Variant** 值，它代表对象的格式代码。

## 说明

**Format** 函数使用的格式代码字符串与 **NumberFormat** 和 **NumberFormatLocal** 属性使用的格式代码字符串不同。

## 示例

```javascript
/*此示例设置图表工作表 Chart1 中图表的第一个数据序列的数据标签的格式代码。*/
function test() {
    let datalabels = Application.Charts.Item("Chart1").ChartObjects(1).Chart.SeriesCollection(1).DataLabels()
    datalabels.NumberFormatLocal = "0%"
}
```

```javascript
/*此示例显示工作表 Sheet1 中第一个图表的第二个数据序列的数据标签的格式。*/
function test() {
    let datalabels = Application.Sheets.Item("Sheet1").ChartObjects(1).Chart.SeriesCollection(2).DataLabels()
    console.log(`数据标签的格式：${datalabels.NumberFormatLocal}`)
}
```
