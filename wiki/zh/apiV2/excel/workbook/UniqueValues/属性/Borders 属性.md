---
title: "Borders (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/UniqueValues/%E5%B1%9E%E6%80%A7/Borders%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: CSq-EwDu
---

# Borders (属性)

返回一个 **[Borders](./../../Borders/Borders 对象.html)** 集合，该集合在条件格式规则的计算结果为 **True** 时指定单元格边框的格式。只读。

## 说明

对于条件格式对象，您只能为单元格的上边框、下边框和侧边框设置属性。

## 示例

```javascript
/*此示例将工作表 Sheet1 中区域 C1:C10 的第一个（UniqueValues）条件格式的右边框设置为蓝色虚线。*/
function test() {
    let uniqueValues = Application.Worksheets.Item("Sheet1").Range("C1:C10").FormatConditions.Item(1)
    let boder = uniqueValues.Borders.Item(xlEdgeRight)
    boder.LineStyle = xlDash
    boder.Color = RGB(0, 0, 255)
}
```

```javascript
/*本示例设置活动工作表上区域 C1:C10 中第二个条件格式边框的线条样式。*/
function test() {
    let uniqueValues = ActiveSheet.Range("C1:C10").FormatConditions.Item(2)
    let boders = uniqueValues.Borders
    boders.LineStyle = xlDouble
}
```
