---
title: "ShapeRange 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/ChartObjects/%E5%B1%9E%E6%80%A7/ShapeRange%20%E5%B1%9E%E6%80%A7.html
---
# ShapeRange (属性)

返回一个 **[ShapeRange](./../../ShapeRange/ShapeRange 对象.html)** 对象，它代表指定的一个或多个对象。只读。

## 示例

```javascript
/*此示例创建一个形状区域，该区域代表第一张工作表上的图表。*/
function test() {
    let sr = Application.Worksheets.Item(1).ChartObjects().ShapeRange
}
```

```javascript
/*此示例选择工作表 Sheet1 上所有图表的形状区域。*/
function test() {
    let chartobjects = Application.Sheets.Item("Sheet1").ChartObjects()
    chartobjects.ShapeRange.Select()
}
```
