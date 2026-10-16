---
title: "Color 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Font/%E5%B1%9E%E6%80%A7/Color%20%E5%B1%9E%E6%80%A7.html
---
# Color (属性)

返回或设置对象的主要颜色，如注释部分中的表格所示。使用 **RGB** 函数可创建颜色值。**Variant** 型，可读写。

## 说明

| 对象 | 对应颜色 |
| --- | --- |
| **Border** | 边框的颜色。 |
| **Borders** | 一个区域的所有四条边的颜色。如果四边不是同一种颜色，则 **Color** 返回的是 0（零）。 |
| **Font** | 字体的颜色。 |
| **Interior** | 单元格底纹的颜色或图形对象的填充颜色。 |
| **Tab** | 选项卡的颜色。 |

## 示例

```javascript
/*此示例对 Chart1 中数值坐标轴的刻度线标志颜色进行设置*/
function test() {
    Application.Charts.Item("Chart1").Axes(xlValue).TickLabels.Font.Color = RGB(0, 255, 0)
}
```

```javascript
/*此示例将工作表Sheet1的单元格 A1:A4 区域中的字体颜色设置为蓝色*/
function test() {
    let font = Application.Worksheets.Item("Sheet1").Range("A1:A4").Font
    font.Color = RGB(0, 0, 255)
}
```
