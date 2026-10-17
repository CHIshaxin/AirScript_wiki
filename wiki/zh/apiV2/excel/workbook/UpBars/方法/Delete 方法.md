---
title: "Delete (方法)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/UpBars/%E6%96%B9%E6%B3%95/Delete%20%E6%96%B9%E6%B3%95.html
updated: 1789700916000
hash: BOp2deyP
---

# Delete (方法)

删除对象。

## 返回值

Variant

## 示例

```javascript
/*本示例删除工作表 Sheet1 中第一个图表的第一个图表组的涨跌柱线。*/
function test() {
    let upbars = Application.Worksheets.Item("Sheet1").ChartObjects(1).Chart.ChartGroups(1).UpBars
    upbars.Delete()
}
```

```javascript
/*本示例删除图表工作表 Chart1 中图表的第一个图表组的涨跌柱线。*/
function test() {
    let upbars = Application.Charts.Item("Chart1").ChartObjects(1).Chart.ChartGroups(1).UpBars
    upbars.Delete()
}
```
