---
title: "IncludeFont (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Style/%E5%B1%9E%E6%80%A7/IncludeFont%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: B7oNAiFz
---

# IncludeFont (属性)

如果样式包含 **Background**、**Bold**、**Color**、**ColorIndex**、**FontStyle**、**Italic**、**Name**、**Size**、**Strikethrough**、**Subscript**、**Superscript**和 **Underline**字体属性，则此属性为 **True**。**Boolean** 类型，可读写。

## 示例

```javascript
/*本示例在 Sheet1 的 A1 单元格样式中加入字体格式。*/
function test() {
    Application.Worksheets.Item("Sheet1").Range("A1").Style.IncludeFont = true
}
```

```javascript
/*本示例显示活动工作表中B2单元格样式是否加入字体格式。*/
function test() {
    let range = ActiveSheet.Range("B2")
    console.log(range.Style.IncludeFont)
}
```
