---
title: "DashStyle 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/LineFormat/%E5%B1%9E%E6%80%A7/DashStyle%20%E5%B1%9E%E6%80%A7.html
---
# DashStyle (属性)

返回或设置指定直线的虚线样式。可以是 **MsoLineDashStyle** 常量之一。可读/写 **Long** 类型。

## 示例

```javascript
/*本示例向第一张工作表中添加蓝色的虚线。*/
function test() {
    let worksheet = Application.Worksheets.Item(1)
    let line = worksheet.Shapes.AddLine(10, 10, 250, 250).Line
    line.DashStyle = msoLineDashDotDot
    line.ForeColor.RGB = RGB(50, 0, 128)
}
```

```javascript
/*本示例向活动工作表添加一条红色实线。*/
function test() {
    let shapes = ActiveSheet.Shapes
    let shape = shapes.AddLine(100, 10, 27, 230)
    shape.Line.DashStyle = msoLineSolid
    shape.Line.ForeColor.RGB = RGB(255, 0, 0)
}
```
