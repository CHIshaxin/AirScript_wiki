---
title: "Caption (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/DataLabel/%E5%B1%9E%E6%80%A7/Caption%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: VXGzCavQ
---

# Caption (属性)

返回或设置一个 **String** 值，它代表数据标签文本。

## 示例

```javascript
/*此示例显示工作表 Sheet1 中第一张图表的第一个数据序列的第二个数据标签的文本。*/
function test() {
    let datalabel = Application.Sheets.Item("Sheet1").ChartObjects(1).Chart.SeriesCollection(1).DataLabels(2)
    console.log(datalabel.Caption)
}
```

```javascript
/*此示例将图表工作表 Chart1 中图表的第一个数据序列的第三个数据标签的文本设置为“三年级”。*/
function test() {
    let datalabel = Application.Charts.Item("Chart1").ChartObjects(1).Chart.SeriesCollection(1).DataLabels(3)
    datalabel.Caption = "三年级"
}
```
