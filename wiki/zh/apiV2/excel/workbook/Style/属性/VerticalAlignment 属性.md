---
title: "VerticalAlignment (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Style/%E5%B1%9E%E6%80%A7/VerticalAlignment%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: BNn0b0q7
---

# VerticalAlignment (属性)

返回或设置一个 **XlVAlign** 值，它代表指定对象的垂直对齐方式。

## 说明

某些常量可能不可用，这取决于所选择或安装的语言支持（例如，美国英语）。

## 示例

```javascript
/*本示例为单元格 B1 赋值，并将单元格垂直对齐方式设置为顶部对齐。*/
function test() {
    Range("B1").Value2 = "示例文本"
    Range("B1").Style.VerticalAlignment = xlVAlignTop
}
```

```javascript
/*本示例为单元格 C1 赋值，并显示单元格垂直对齐方式是否为居中对齐。*/
function test() {
    Range("C1").Value2 = "示例文本"
    console.log(Range("C1").Style.VerticalAlignment == xlVAlignCenter)
}
```
