---
title: "Delete (方法)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/HiLoLines/%E6%96%B9%E6%B3%95/Delete%20%E6%96%B9%E6%B3%95.html
updated: 1789700916000
hash: BCgH_bav
---

# Delete (方法)

删除对象。

## 返回值

Variant

## 示例

```javascript
/*本示例判断如果工作表 Sheet1 上第一个图表中第一个图表组有高低点连线，则将高低点连线删除。*/
function test() {
    let chartGroup = Worksheets.Item("Sheet1").ChartObjects(1).Chart.ChartGroups(1)
    if (chartGroup.HasHiLoLines) {
        chartGroup.HiLoLines.Delete()
    }
}
```

```javascript
/*本示例将图表工作表 Chart1 中第一个图表组的高低点连线删除。*/
function test() {
    let chartGroup = Application.Charts.Item("Chart1").ChartObjects(1).Chart.ChartGroups(1)
    chartGroup.HiLoLines.Delete()
}
```
