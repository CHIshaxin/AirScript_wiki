---
title: "Font 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/DataTable/%E5%B1%9E%E6%80%A7/Font%20%E5%B1%9E%E6%80%A7.html
---
# Font (属性)

返回一个 **[Font](./../../Font/Font 对象.html)** 对象，它代表指定对象的字体。

## 示例

```javascript
/*本示例将工作表 Sheet1 中第一张图表的模拟运算表中的字体设置为蓝色粗体。*/
function test() {
    let datatable = Application.Worksheets.Item("Sheet1").ChartObjects(1).Chart.DataTable
    datatable.Font.Bold = true
    datatable.Font.ColorIndex = 5
}
```

```javascript
/*本示例将图表工作表 Chart1 中图表的模拟运算表中的字体设置为加粗和倾斜，并将字体字号设置为 12 磅。*/
function test() {
    let datatable = Application.Charts.Item("Chart1").ChartObjects(1).Chart.DataTable
    datatable.Font.FontStyle = "Bold Italic"
    datatable.Font.Size = 12
}
```
