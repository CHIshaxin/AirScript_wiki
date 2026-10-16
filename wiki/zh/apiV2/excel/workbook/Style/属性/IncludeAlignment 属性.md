---
title: "IncludeAlignment (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Style/%E5%B1%9E%E6%80%A7/IncludeAlignment%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: Bqsb5ItK
---

# IncludeAlignment (属性)

如果样式包含 **AddIndent**、**HorizontalAlignment**、**VerticalAlignment**、**WrapText**、**IndentLevel** 和 **Orientation** 属性，则该属性值为 **True**。**Boolean** 类型，可读写。

## 示例

```javascript
/*本示例在 Sheet1 的 A1 单元格样式中加入对齐格式。*/
function test() {
    Application.Worksheets.Item("Sheet1").Range("A1").Style.IncludeAlignment = true
}
```

```javascript
/*本示例在活动工作表中B1单元格添加文本，并显示单元格样式是否加入对齐格式。*/
function test() {
    let range = ActiveSheet.Range("B1")
    range.Value2 = "这是示例文本"
    console.log(range.Style.IncludeAlignment)
}
```
