---
title: "Color (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Border/%E5%B1%9E%E6%80%A7/Color%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: J8O-j3C5
---

# Color (属性)

返回或设置对象的主要颜色，如注释部分中的表格所示。**Variant** 型，可读写。

## 说明

| 对象 | 对应颜色 |
|---|---|
| **Border** | 边框的颜色。 |
| **Borders** | 一个区域的所有四条边的颜色。如果四边不是同一种颜色，则 **Color** 返回的是 0（零）。 |
| **Font** | 字体的颜色。 |
| **Interior** | 单元格底纹的颜色或图形对象的填充颜色。 |
| **Tab** | 选项卡的颜色。 |

## 示例

```javascript
/*此示例对 Chart1 中数值坐标轴的边框颜色进行设置。*/
function test() {
    Application.Charts.Item("Chart1").Axes(xlValue).Border.Color = RGB(0, 255, 0)
}
```

```javascript
/*本示例设置 H5 单元格左边框的颜色为红色。*/
function test() {
    Range("H5").Borders.Item(xlEdgeLeft).Color = RGB(255, 0, 0)
}
```
