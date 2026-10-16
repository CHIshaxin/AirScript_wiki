---
title: "Color (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Interior/%E5%B1%9E%E6%80%A7/Color%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: DpPMArNY
---

# Color (属性)

返回或设置对象的主要颜色，如注释部分中的表格所示。使用 **RGB** 函数可创建颜色值。**Variant** 型，可读写。

## 说明

| **Border** | 边框的颜色。 |
| --- | --- |
| **Borders** | 一个区域的所有四条边的颜色。如果四边不是同一种颜色，则 **Color** 返回的是 0（零）。 |
| **Font** | 字体的颜色。 |
| **Interior** | 单元格底纹的颜色或图形对象的填充颜色。 |
| **Tab** | 选项卡的颜色。 |

## 示例

```javascript
/*此示例将活动工作表B2单元格背景颜色设置为黄色。*/
function test() {
    ActiveSheet.Range("B2").Interior.Color = RGB(255, 255, 0)
}
```

```javascript
/*本示例将 Sheet1 中 A1:C4 区域单元格背景颜色设置为红色。*/
function test() {
    Application.Worksheets.Item("Sheet1").Range("A1:C4").Interior.Color = RGB(255, 0, 0)
}
```
