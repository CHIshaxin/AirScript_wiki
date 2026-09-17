---
title: "Delete 方法"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Gridlines/%E6%96%B9%E6%B3%95/Delete%20%E6%96%B9%E6%B3%95.html
---
# Delete (方法)

删除对象。

## 返回值

Variant

## 示例

```javascript
/*本示例删除图表工作表 Chart1 数值轴的主要网格线。*/
function test() {
    let axis = Application.Charts.Item("Chart1").ChartObjects(1).Chart.Axes(xlValue)
    axis.MajorGridlines.Delete()
}
```

```javascript
/*本示例判断如果工作表 Sheet1 中第一个图表类别轴有主要网格线，则将主要网格线删除。*/
function test() {
    let axis = Worksheets.Item("Sheet1").ChartObjects(1).Chart.Axes(xlCategory)
    if (axis.HasMajorGridlines) {
        axis.MajorGridlines.Delete()
    }
}
```
