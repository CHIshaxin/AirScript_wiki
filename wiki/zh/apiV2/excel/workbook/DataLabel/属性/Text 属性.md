---
title: "Text (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/DataLabel/%E5%B1%9E%E6%80%A7/Text%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: BJuo0pyR
---

# Text (属性)

返回或设置指定对象中的文本。**String** 型，可读写。

## 示例

```javascript
/*此示例显示图表工作表 Chart1 中图表的第二个系列的第三个数据标签的文本。*/
function test() {
    let datalabel = Application.Charts.Item("Chart1").ChartObjects(1).Chart.SeriesCollection(2).DataLabels(3)
    console.log(datalabel.Text)
}
```

```javascript
/*此示例将工作表 Sheet1 中第一张图表的第一个系列的第一个数据标签的文本设置为“软工2班”。*/
function test() {
    let datalabel = Application.Sheets.Item("Sheet1").ChartObjects(1).Chart.SeriesCollection(1).DataLabels(1)
    datalabel.Text = "软工2班"
}
```
