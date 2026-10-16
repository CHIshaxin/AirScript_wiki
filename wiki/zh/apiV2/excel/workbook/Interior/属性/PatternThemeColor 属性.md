---
title: "PatternThemeColor (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Interior/%E5%B1%9E%E6%80%A7/PatternThemeColor%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: BA-DrWaG
---

# PatternThemeColor (属性)

返回或设置 **[Interior](./../../Interior/Interior 对象.html)** 对象的主题颜色图案。可读/写 **Variant** 类型。

## 示例

```javascript
/*本示例为活动工作表中 A1:B2 区域单元格添加网格图案，并将图案主题颜色设置为 xlThemeColorAccent1。*/
function test() {
    let interior = ActiveSheet.Range("A1:B2").Interior
    interior.Pattern = xlPatternGrid
    interior.PatternThemeColor = xlThemeColorAccent1
}
```

```javascript
/*本示例显示 Sheet1 中 C3 单元格中图案主题颜色是否为 xlThemeColorLight1。*/
function test() {
    let interior = Application.Worksheets.Item("Sheet1").Range("C3").Interior
    console.log(interior.PatternThemeColor == xlThemeColorLight1)
}
```
