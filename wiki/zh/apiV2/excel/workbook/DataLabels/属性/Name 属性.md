---
title: "Name (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/DataLabels/%E5%B1%9E%E6%80%A7/Name%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: BaQ3-hn4
---

# Name (属性)

返回一个 **String** 值，它代表对象的名称。

## 示例

```javascript
/*此示例显示工作表 Sheet1 中第一个图表的第一个数据序列的数据标签对象的名称。*/
function test() {
    let datalabels = Application.Sheets.Item("Sheet1").ChartObjects(1).Chart.SeriesCollection(1).DataLabels()
    console.log(datalabels.Name)
}
```

```javascript
/*此示例显示图表工作表 Chart1 中图表的第二个数据序列的数据标签对象的名称。*/
function test() {
    let datalabels = Application.Charts.Item("Chart1").ChartObjects(1).Chart.SeriesCollection(2).DataLabels()
    console.log(`数据标签对象的名称：${datalabels.Name}`)
}
```
