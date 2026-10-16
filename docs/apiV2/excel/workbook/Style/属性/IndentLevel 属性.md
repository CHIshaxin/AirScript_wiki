---
title: "IndentLevel 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Style/%E5%B1%9E%E6%80%A7/IndentLevel%20%E5%B1%9E%E6%80%A7.html
---
# IndentLevel (属性)

返回或设置一个 **Long** 值，它代表样式的缩进量。

## 说明

使用此属性将缩进量设为小于 0 或者大于 15 的数字，将导致发生错误。

## 示例

```javascript
/*本示例显示A1单元格的缩进量。*/
function test() {
    let range = Application.Worksheets.Item(1).Range("A1")
    range.Value2 = "这是示例文本"
    console.log(range.Style.IndentLevel)
}
```

```javascript
/*本示例将B2单元格的缩进量设置为5。*/
function test() {
    let range = ActiveSheet.Range("B2")
    range.Style.IndentLevel = 5
}
```
