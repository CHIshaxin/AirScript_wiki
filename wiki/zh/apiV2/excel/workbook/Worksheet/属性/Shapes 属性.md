---
title: "Shapes (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Worksheet/%E5%B1%9E%E6%80%A7/Shapes%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: BWx8Imt4
---

# Shapes (属性)

返回一个 Shapes集合，它代表工作表上的所有形状。只读。

## 示例

```javascript
/*此示例向第一张工作表中添加兰色的虚线。*/
function test() {
    let Wshapes = Application.Worksheets.Item(1).Shapes.AddLine(10, 10, 250, 250).Line
    Wshapes.DashStyle = msoLineDashDotDot
    Wshapes.ForeColor.RGB = RGB(50, 0, 128)
}
```

```javascript
/*本示例选中活动工作表上的所有形状。*/
function test() {
    Application.ActiveSheet.Shapes.SelectAll()
}
```
