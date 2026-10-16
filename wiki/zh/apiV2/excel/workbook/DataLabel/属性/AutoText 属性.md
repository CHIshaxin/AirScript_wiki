---
title: "AutoText (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/DataLabel/%E5%B1%9E%E6%80%A7/AutoText%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: mTxCrTq9
---

# AutoText (属性)

如果对象会根据内容自动生成合适的文字，则为 **True**。**Boolean** 类型，可读写。

## 示例

```javascript
/*此示例对 Chart1 的第一个数据序列的第一个数据标签进行设置，自动生成合适的文字。*/
function test() {
    let datalabel = Application.Charts.Item("Chart1").ChartObjects(1).Chart.SeriesCollection(1).DataLabels(1)
    datalabel.AutoText = true
}
```

```javascript
/*此示例显示工作表 Sheet1 中第一张图表的第二个数据序列的第二个数据标签是否自动生成合适的文字。*/
function test() {
    let datalabel = Application.Sheets.Item("Sheet1").ChartObjects(1).Chart.SeriesCollection(2).DataLabels(2)
    console.log(datalabel.AutoText)
}
```
