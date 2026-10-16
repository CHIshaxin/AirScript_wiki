---
title: "Orientation (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/DataLabel/%E5%B1%9E%E6%80%A7/Orientation%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: BGs5I5eP
---

# Orientation (属性)

返回或设置一个 **Variant** 值，它代表文本方向。

## 说明

此属性的值可以设置为 -90 到 90 度的整数值，或设置为 **[XlOrientation](./../../枚举/XlOrientation 枚举.html)** 常量之一。

## 示例

```javascript
/*此示例显示工作表 Sheet1 中第一个图表的第一个数据序列的第一个数据标签的文字方向是否为水平排列。*/
function test() {
    let datalabel = Application.Sheets.Item("Sheet1").ChartObjects(1).Chart.SeriesCollection(1).DataLabels(1)
    console.log(datalabel.Orientation == xlHorizontal)
}
```

```javascript
/*此示例将图表工作表 Chart1 中图表的第一个数据序列的第二个数据标签的文字方向设置为向下排列。*/
function test() {
    let datalabel = Application.Charts.Item("Chart1").ChartObjects(1).Chart.SeriesCollection(1).DataLabels(2)
    datalabel.Orientation = xlDownward
}
```
