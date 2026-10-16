---
title: "PrefixCharacter (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Range/%E5%B1%9E%E6%80%A7/PrefixCharacter%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: B38WV2lA
---

# PrefixCharacter (属性)

返回单元格的前缀字符。**Variant** 类型，只读。

## 说明

如果 **TransitionNavigKeys** 属性为 **False**，则对于文本标签，该前缀字符将是“'” 或空白。如果 **TransitionNavigKeys** 属性为 **True**，则对于左对齐标签，该字符将是“'”，对于右对齐标签，是“"”，对于居中对齐标签，是“^”，对于重复标签，是“\”或空白。

## 示例

```javascript
/*本示例显示 sheet1 中单元格 A1 的前缀字符。*/
function test() {
    console.log("The prefix character is " + Worksheets.Item("Sheet1").Range("A1").PrefixCharacter)
}
```

```javascript
/*本示例将 C3 单元格的值设置为 'test，并显示 C3 单元格的前缀字符*/
function test() {
    Range("C3").Value2 = "'test1"
    console.log(Range("C3").PrefixCharacter)
}
```
