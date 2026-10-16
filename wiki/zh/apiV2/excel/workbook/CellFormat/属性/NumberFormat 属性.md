---
title: "NumberFormat (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/CellFormat/%E5%B1%9E%E6%80%A7/NumberFormat%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: 8p9ZUhZf
---

# NumberFormat (属性)

返回或设置一个 **Variant** 值，它代表对象的格式代码。

## 说明

如果指定区域中的所有单元格包含不同的数字格式，则此属性返回 **Null**。

格式代码与**“设置单元格格式”**对话框中的**“格式代码”**选项是同一个字符串。**Format** 函数使用的格式代码字符串与 **NumberFormat** 和 **NumberFormatLocal** 属性使用的格式代码字符串不同。

## 示例

```javascript
/*此示例显示查找条件的格式代码。*/
function test() {
    console.log(Application.FindFormat.NumberFormat)
}
```

```javascript
/*此示例设置替换条件的格式代码。*/
function test() {
    Application.ReplaceFormat.NumberFormat = "0.00"
}
```
