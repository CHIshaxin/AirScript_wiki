---
title: "LineFormat (对象)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/LineFormat/LineFormat%20%E5%AF%B9%E8%B1%A1.html
updated: 1789700916000
hash: o3jYN9Ev
---

# LineFormat (对象)

代表线条和箭头格式。

## 说明

对于线条， **[LineFormat](./../LineFormat/LineFormat 对象.html)** 对象包含该线条自身的格式信息；对于有边界的形状，该对象包含形状的边界的格式信息。

使用 **Line**属性可返回一个 **[LineFormat](./../LineFormat/LineFormat 对象.html)** 对象。

## 示例

```javascript
/*下例给第一张工作表添加一条蓝色虚线。在该线的起点有一个短而窄的椭圆，在该线的终点有一个长而宽的三角形。*/
function test() {
    let worksheet = Application.Worksheets.Item(1)
    let line = worksheet.Shapes.AddLine(100, 100, 200, 300).Line
    line.DashStyle = msoLineDashDotDot
    line.ForeColor.RGB = RGB(50, 0, 128)
    line.BeginArrowheadLength = msoArrowheadShort
    line.BeginArrowheadStyle = msoArrowheadOval
    line.BeginArrowheadWidth = msoArrowheadNarrow
    line.EndArrowheadLength = msoArrowheadLong
    line.EndArrowheadStyle = msoArrowheadTriangle
    line.EndArrowheadWidth = msoArrowheadWide
}
```

```javascript
/*本示例显示活动工作表的第一个形状线条的样式是否为双细线。*/
function test() {
    let line = ActiveSheet.Shapes.Item(1).Line
    console.log(line.Style == msoLineThinThin)
}
```
