---
title: "WrapText (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/CellFormat/%E5%B1%9E%E6%80%A7/WrapText%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: D8ywl5gf
---

# WrapText (属性)

返回或设置一个 **Variant** 值，它指明 ET 是否为对象中的文本自动换行。

## 说明

如果指定区域内所有单元格中的文本都自动换行，此属性将返回 **True**；如果指定区域内所有单元格中的文本都不自动换行，则返回 **False**；如果指定区域内有些单元格中的文本自动换行，而另一些单元格中的文本不自动换行，则返回 **Null**。

ET 会在必要的时候改变区域的行高以容纳其中的文字。

## 示例

```javascript
/*此示例设置替换条件中文本不自动换行。*/
function test() {
    Application.ReplaceFormat.WrapText = false
}
```

```javascript
/*此示例显示查找条件中文本是否自动换行。*/
function test() {
    console.log(Application.FindFormat.WrapText)
}
```
