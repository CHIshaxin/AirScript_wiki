---
title: "FontStyle 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Font/%E5%B1%9E%E6%80%A7/FontStyle%20%E5%B1%9E%E6%80%A7.html
---
# FontStyle (属性)

返回或设置字体样式。**String** 类型，可读写。

## 说明

更改该属性可能会影响其他 **Font** 属性（例如 **Bold** 和 **Italic**）。

## 示例

```javascript
/*本示例将 Sheet1 中 A1 单元格的字体样式设置为加粗和倾斜*/
function test() {
    Application.Worksheets.Item("Sheet1").Range("A1").Font.FontStyle = "Bold Italic"
}
```

```javascript
/*此示例显示C1单元格的字体样式*/
function test() {
    let font = Application.Worksheets.Item(1).Range("C1").Font
    console.log(`字体样式为 ${font.FontStyle}`)
}
```
