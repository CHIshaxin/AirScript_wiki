---
title: "Height 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/DataLabel/%E5%B1%9E%E6%80%A7/Height%20%E5%B1%9E%E6%80%A7.html
---
# Height (属性)

返回对象的高度（以磅为单位）。只读。

## 示例

```javascript
/*此示例显示工作表 Sheet1 中第一张图表的第二个数据序列的第一个数据标签的高度。*/
function test() {
    let datalabel = Application.Sheets.Item("Sheet1").ChartObjects(1).Chart.SeriesCollection(2).DataLabels(1)
    console.log(`数据标签的高度：${datalabel.Height} 磅`)
}
```

```javascript
/*此示例将图表工作表 Chart1 中图表的第一个数据序列的第三个数据标签的文本设置为“三年级”，并显示该数据标签的高度。*/
function test() {
    let datalabel = Application.Charts.Item("Chart1").ChartObjects(1).Chart.SeriesCollection(1).DataLabels(3)
    datalabel.Caption = "三年级"
    console.log(datalabel.Height)
}
```
