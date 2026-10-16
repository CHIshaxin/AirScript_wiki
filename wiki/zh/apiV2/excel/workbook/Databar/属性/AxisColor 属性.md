---
title: "AxisColor (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Databar/%E5%B1%9E%E6%80%A7/AxisColor%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: COCSoshY
---

# AxisColor (属性)

返回条件格式为数据条的单元格的坐标轴颜色。只读。

## 返回值

FormatColor

## 示例

```javascript
/*本示例设置工作表 Sheet1 上区域 A1:A10 中第一个为数据条的条件格式的单元格的坐标轴颜色。*/
function test() {
    let databar = Application.Worksheets.Item("Sheet1").Range("A1:A10").FormatConditions.Item(1)
    databar.AxisColor.ColorIndex = 3
}
```

```javascript
/*本示例设置工作表 Sheet1 上区域 A1:A10 中第一个为数据条的条件格式的单元格的坐标轴的主题颜色。*/
function test() {
    let databar = Application.Worksheets.Item("Sheet1").Range("A1:A10").FormatConditions.Item(1)
    databar.AxisColor.ThemeColor = xlThemeColorAccent1
}
```
