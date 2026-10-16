---
title: "ShapeRange (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/ChartObject/%E5%B1%9E%E6%80%A7/ShapeRange%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: ppmdW4zB
---

# ShapeRange (属性)

返回一个 **[ShapeRange](./../../ShapeRange/ShapeRange 对象.html)** 对象，它代表指定的一个或多个对象。只读。

## 示例

```javascript
/*此示例创建一个形状区域，该区域代表第一张工作表上的图表。*/
function test() {
    let shprange = Application.Worksheets.Item(1).ChartObjects().ShapeRange
}
```

```javascript
/*此示例删除工作表 Sheet1 中图表 Chart 2 的形状区域。*/
function test() {
    let chartobject = Application.Worksheets.Item("Sheet1").ChartObjects("Chart 2")
    chartobject.ShapeRange.Delete()
}
```
