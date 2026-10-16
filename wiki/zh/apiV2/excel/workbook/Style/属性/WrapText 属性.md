---
title: "WrapText (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Style/%E5%B1%9E%E6%80%A7/WrapText%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: DSRcjo02
---

# WrapText (属性)

返回或设置一个 **Boolean** 值，它指明 ET 是否为对象中的文本自动换行。

## 示例

```javascript
/*本示例为单元格 B1 赋值，并设置单元格的文本可自动换行。*/
function test() {
    Range("B1").Value2 = "这是示例文本"
    Range("B1").Style.WrapText = true
}
```

```javascript
/*本示例为单元格 C1 赋值，并显示单元格的文本是否可以自动换行。*/
function test() {
    Range("C1").Value2 = "这是示例文本"
    console.log(Range("C1").Style.WrapText)
}
```
