---
title: "Delete 方法"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/DataLabels/%E6%96%B9%E6%B3%95/Delete%20%E6%96%B9%E6%B3%95.html
---
# Delete (方法)

删除对象。

## 返回值

Variant

## 示例

```javascript
/*此示例删除图表工作表 Chart1 中图表的第一个数据序列的数据标签。*/
function test() {
    let datalabels = Application.Charts.Item("Chart1").ChartObjects(1).Chart.SeriesCollection(1).DataLabels()
    datalabels.Delete()
}
```

```javascript
/*此示例判断如果工作表 Sheet1 中第一张图表的第二个数据序列的数据标签的数量为 4 ，则删除该数据标签。*/
function test() {
    let datalabels = Application.Sheets.Item("Sheet1").ChartObjects(1).Chart.SeriesCollection(2).DataLabels()
    if (datalabels.Count == 4) {
        datalabels.Delete()
    }
}
```
