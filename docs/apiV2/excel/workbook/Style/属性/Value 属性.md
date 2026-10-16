---
title: "Value 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Style/%E5%B1%9E%E6%80%A7/Value%20%E5%B1%9E%E6%80%A7.html
---
# Value (属性)

返回一个 **String** 值，它代表指定样式的名称。

## 示例

```javascript
/*本示例为A1单元格添加文本，并显示A1单元格样式的名称。*/
function test() {
    Range("A1").Value2 = "示例文本"
    console.log(Range("A1").Style.Value)
}
```

```javascript
/*本示例显示B1单元格样式的名称是否为“Percent”。*/
function test() {
    console.log(Range("B1").Style.Value == "Percent")
}
```
