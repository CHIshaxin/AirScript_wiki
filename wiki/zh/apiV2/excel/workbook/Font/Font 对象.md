---
title: "Font (对象)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Font/Font%20%E5%AF%B9%E8%B1%A1.html
updated: 1789700916000
hash: CS2vvHIG
---

# Font (对象)

包含对象的字体属性（字体名称、字号、颜色等等）。

## 说明

如果不想将单元格中的文本或图形设为相同的格式，则使用 **Characters** 属性返回文本的子集。

使用 **Font** 属性可返回 **[Font](./../Font/Font 对象.html)** 对象。

## 示例

```javascript
/*本示例将工作表 Sheet1中单元格 A1:C5 区域中的字体设为加粗*/
function test() {
    Application.Worksheets.Item("Sheet1").Range("A1:C5").Font.Bold = true
}
```

```javascript
/*此示例显示第一个工作表中D1单元格的字体样式*/
function test() {
    let font = Application.Worksheets.Item(1).Range("D1").Font
    console.log(`字体样式为 ${font.FontStyle}`)
}
```

``
