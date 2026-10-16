---
title: "IncludeProtection (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Style/%E5%B1%9E%E6%80%A7/IncludeProtection%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: DzuXZ8Rt
---

# IncludeProtection (属性)

如果指定样式中包含 **FormulaHidden** 和 **Locked** 保护属性，则该属性值为 **True**。**Boolean** 类型，可读写。

## 示例

```javascript
/*本示例在 Sheet1 的 A1 单元格样式中加入保护格式。*/
function test() {
    Application.Worksheets.Item("Sheet1").Range("A1").Style.IncludeProtection = true
}
```

```javascript
/*本示例显示活动工作表中B2单元格样式是否加入保护格式。*/
function test() {
    let range = ActiveSheet.Range("B2")
    console.log(range.Style.IncludeProtection)
}
```
