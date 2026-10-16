---
title: "NumberFormatLocal (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/CellFormat/%E5%B1%9E%E6%80%A7/NumberFormatLocal%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: Biq1qo5L
---

# NumberFormatLocal (属性)

以采用用户语言字符串的形式返回或设置一个 **Variant** 值，它代表对象的格式代码。

## 说明

**Format** 函数使用的格式代码字符串与 **NumberFormat** 和 **NumberFormatLocal** 属性使用的格式代码字符串不同。

## 示例

```javascript
/*此示例设置查找条件的格式代码。*/
function test() {
    Application.FindFormat.NumberFormatLocal = "0.00%"
}
```

```javascript
/*此示例显示替换条件的格式代码。*/
function test() {
    console.log(Application.ReplaceFormat.NumberFormatLocal)
}
```
