---
title: "Format (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/DataTable/%E5%B1%9E%E6%80%A7/Format%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: Ndr85_fQ
---

# Format (属性)

返回 **[ChartFormat](./../../ChartFormat/ChartFormat 对象.html)** 对象。只读。

## 示例

```javascript
/*本示例将图表工作表 Chart1 中图表的模拟运算表的前景色设置为青色。*/
function test() {
    let datatable = Application.Charts.Item("Chart1").ChartObjects(1).Chart.DataTable
    datatable.Format.Fill.ForeColor.RGB = RGB(0, 255, 255)
}
```

```javascript
/*本示例将工作表 Sheet1 中第一张图表的模拟运算表的线条设置为红色的点划线。*/
function test() {
    let datatable = Application.Sheets.Item("Sheet1").ChartObjects(1).Chart.DataTable
    datatable.Format.Line.DashStyle = msoLineDashDot
    datatable.Format.Line.ForeColor.RGB = RGB(255, 0, 0)
}
```
