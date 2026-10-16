---
title: "Select (方法)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/DataLabel/%E6%96%B9%E6%B3%95/Select%20%E6%96%B9%E6%B3%95.html
updated: 1789700916000
hash: DrTXYBA8
---

# Select (方法)

选择对象。

## 返回值

Variant

## 示例

```javascript
/*此示例选中图表工作表 Chart1 中图表的第一个数据序列的第一个数据标签。*/
function test() {
    let datalabel = Application.Charts.Item("Chart1").ChartObjects(1).Chart.SeriesCollection(1).DataLabels(1)
    datalabel.Select()
}
```

```javascript
/*此示例判断如果工作表 Sheet1 中第一个图表的第二个数据序列的第一个数据标签的文本为“三年级”，则选中该数据标签。*/
function test() {
    let datalabel = Application.Sheets.Item("Sheet1").ChartObjects(1).Chart.SeriesCollection(2).DataLabels(1)
    if (datalabel.Caption == "三年级") {
        datalabel.Select()
    }
}
```
