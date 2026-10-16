---
title: "Bold (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Font/%E5%B1%9E%E6%80%A7/Bold%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: CozUmXo_
---

# Bold (属性)

如果字体格式为加粗，则该属性值为 **True**。**Variant** 类型，可读写。

## 示例

```javascript
/*本示例将工作表 Sheet1中单元格 A1:A5 区域中的字体设为加粗*/
function test() {
    Application.Worksheets.Item("Sheet1").Range("A1:A5").Font.Bold = true
}
```

```javascript
/*本示例显示第一个工作表中的单元格A1字体是否为加粗。*/
function test() {
    let bold = Application.Worksheets.Item(1).Range("A1").Font.Bold
    console.log(bold)
}
```
