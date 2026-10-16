---
title: "IncludeNumber (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Style/%E5%B1%9E%E6%80%A7/IncludeNumber%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: BslY5ezw
---

# IncludeNumber (属性)

如果样式中包含 **NumberFormat** 属性，则该属性值为 **True**。**Boolean** 类型，可读写。

## 示例

```javascript
/*本示例在 Sheet1 的 A1 单元格样式中加入数字格式。*/
function test() {
    Application.Worksheets.Item("Sheet1").Range("A1").Style.IncludeNumber = true
}
```

```javascript
/*本示例显示活动工作表中D2单元格样式是否加入数字格式。*/
function test() {
    let range = ActiveSheet.Range("D2")
    console.log(range.Style.IncludeNumber)
}
```
