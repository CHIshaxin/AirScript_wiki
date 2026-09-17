---
title: "Width 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/ChartObjects/%E5%B1%9E%E6%80%A7/Width%20%E5%B1%9E%E6%80%A7.html
---
# Width (属性)

返回或设置一个 **Double** 值，它代表对象的宽度（以磅为单位）。

## 示例

```javascript
/*此示例显示图表工作表 Chart1 上图表的宽度。*/
function test() {
    let chartobjects = Application.Charts.Item("Chart1").ChartObjects()
    console.log(chartobjects.Width)
}
```

```javascript
/*此示例设置工作表 Sheet1 上所有图表的左边界与 B 列的左边界对齐，并将所有图表的宽度设置为 A1:F1 的距离。*/
function test() {
    let sheet = Application.Worksheets.Item("Sheet1")
    sheet.ChartObjects().Left = sheet.Columns.Item("B").Left
    sheet.ChartObjects().Width = sheet.Range("A1:F1").Width
}
```
