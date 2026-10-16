---
title: "Delete (方法)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/DropLines/%E6%96%B9%E6%B3%95/Delete%20%E6%96%B9%E6%B3%95.html
updated: 1789700916000
hash: Q1FxhJ3w
---

# Delete (方法)

删除对象。

## 返回值

Variant

## 示例

```javascript
/*本示例删除工作表 Sheet1 中第一个图表的第一个图表组的垂直线。*/
function test() {
    let droplines = Application.Worksheets.Item("Sheet1").ChartObjects(1).Chart.ChartGroups(1).DropLines
    droplines.Delete()
}
```

```javascript
/*本示例判断如果图表工作表 Chart1 中图表的第一个图表组的垂直线的颜色为红色，则删除该垂直线。*/
function test() {
    let droplines = Application.Charts.Item("Chart1").ChartObjects(1).Chart.ChartGroups(1).DropLines
    if (droplines.Border.ColorIndex == 3) {
        droplines.Delete()
    }
}
```
