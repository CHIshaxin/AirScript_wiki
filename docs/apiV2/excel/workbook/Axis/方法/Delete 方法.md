---
title: "Delete 方法"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Axis/%E6%96%B9%E6%B3%95/Delete%20%E6%96%B9%E6%B3%95.html
---
# Delete (方法)

删除对象。

## 返回值

Variant

## 示例

```javascript
/*本示例删除工作表 Sheet1 中第一张图表的数值坐标轴。*/
function test() {
    let axis = Application.Sheets.Item("Sheet1").ChartObjects(1).Chart.Axes(xlValue)
    axis.Delete()
}
```

```javascript
/*本示例删除图表工作表 Chart1 中图表的分类坐标轴。*/
function test() {
    let axis = Application.Charts.Item("Chart1").ChartObjects(1).Chart.Axes(xlCategory)
    axis.Delete()
}
```
