---
title: "Italic (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Font/%E5%B1%9E%E6%80%A7/Italic%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: CZ0EdoEt
---

# Italic (属性)

如果字体样式为倾斜，则该属性值为 **True**。**Boolean** 类型，可读写。

## 示例

```javascript
/*本示例将 Sheet1 中 A1:A5 区域的字形设置为倾斜*/
function test() {
    Application.Worksheets.Item("Sheet1").Range("A1:A5").Font.Italic = true
}
```

```javascript
/*本示例显示第一个工作表中单元格D2的字形是否为倾斜。*/
function test() {
    let font = Application.Worksheets.Item(1).Range("D2").Font
    console.log(font.Italic)
}
```
