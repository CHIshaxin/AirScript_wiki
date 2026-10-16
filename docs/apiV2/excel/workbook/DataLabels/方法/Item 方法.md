---
title: "Item 方法"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/DataLabels/%E6%96%B9%E6%B3%95/Item%20%E6%96%B9%E6%B3%95.html
---
# Item (方法)

从集合中返回一个对象。

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| Index | any | 必填 | 对象的索引号。 |

## 返回值

包含在集合中的一个 DataLabel 对象。

## 示例

```javascript
/*此示例设置工作表 Sheet1 中第一张图表的第一个数据系列的第三个数据标签的数字格式。*/
function test() {
    let datalabel = Application.Worksheets.Item("Sheet1").ChartObjects(1).Chart.SeriesCollection(1).DataLabels().Item(3)
    datalabel.NumberFormat = "0.000"
}
```

```javascript
/*此示例将图表工作表 Chart1 中图表的第二个数据序列的第三个数据标签的文本设置为“三年级”。*/
function test() {
    let datalabels = Application.Charts.Item("Chart1").ChartObjects(1).Chart.SeriesCollection(2).DataLabels().Item(3)
    datalabels.Caption = "三年级"
}
```
