---
title: "Background (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Font/%E5%B1%9E%E6%80%A7/Background%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: CUZy0RXK
---

# Background (属性)

返回或设置图表中使用的文本的背景类型。**Variant** 类型，可读写。可将其设置为 **[XlBackground](./../../枚举/XlBackground 枚举.html)** 常量之一。

## 说明

| XlBackground 可为以下常量之一。 | 文本背景描述 |
| --- | --- |
| **xlBackgroundAutomatic** | 字体背景将自动更改文本周围的背景区域颜色，以使在文本下方元素所应用的颜色之上显示的文本达到最佳效果 |
| **xlBackgroundOpaque** | 如果文本颜色与文本下的填充颜色非常接近或者为同一颜色，则字体背景会设置为空，以便不显示文本 |
| **xlBackgroundTransparent** | 字体背景设置为透明，以便当文本颜色接近文本下方的颜色时，文本背景不会更改 |

## 示例

```javascript
/*本示例为第一张工作表上的第一张嵌入图表添加图表标题，然后设置标题的字体大小和背景类型。本示例假定图表在第一张工作表上。*/
function test() {
    let chart = Application.Worksheets.Item(1).ChartObjects(1).Chart
    chart.HasTitle = true
    chart.ChartTitle.Text = "Rainfall Totals by Month"
    let font = chart.ChartTitle.Font
    font.Size = 10
    font.Background = xlBackgroundTransparent
}
```

```javascript
/*本示例将 A2 单元格的背景类型设置为 xlBackgroundOpaque 。*/
function test() {
    Range("A2").Font.Background = xlBackgroundOpaque
}
```
