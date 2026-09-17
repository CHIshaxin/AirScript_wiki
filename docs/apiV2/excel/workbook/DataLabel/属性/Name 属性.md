---
title: "Name 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/DataLabel/%E5%B1%9E%E6%80%A7/Name%20%E5%B1%9E%E6%80%A7.html
---
# Name (属性)

返回一个 **String** 值，它代表对象的名称。

## 示例

```javascript
/*此示例显示工作表 Sheet1 中第一个图表的第二个数据序列的第一个数据标签对象的名称。*/
function test() {
    let datalabel = Application.Sheets.Item("Sheet1").ChartObjects(1).Chart.SeriesCollection(2).DataLabels(1)
    console.log(datalabel.Name)
}
```

```javascript
/*此示例显示图表工作表 Chart1 中图表的第三个数据序列的第一个数据标签对象的名称。*/
function test() {
    let datalabel = Application.Charts.Item("Chart1").ChartObjects(1).Chart.SeriesCollection(3).DataLabels(1)
    console.log(`数据标签对象的名称：${datalabel.Name}`)
}
```
