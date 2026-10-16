---
title: "IncludeBorder 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Style/%E5%B1%9E%E6%80%A7/IncludeBorder%20%E5%B1%9E%E6%80%A7.html
---
# IncludeBorder (属性)

如果指定样式中包含 **Color**、**ColorIndex**、**LineStyle** 和 **Weight** 边框属性，则该属性值为 **True**。**Boolean** 类型，可读写。

## 示例

```javascript
/*本示例在 Sheet1 的 A1 单元格样式中加入边框格式。*/
function test() {
    Application.Worksheets.Item("Sheet1").Range("A1").Style.IncludeBorder = true
}
```

```javascript
/*本示例显示活动工作表中B2单元格样式是否加入边框格式。*/
function test() {
    let range = ActiveSheet.Range("B2")
    console.log(range.Style.IncludeBorder)
}
```
