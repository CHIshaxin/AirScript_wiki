---
title: "HorizontalAlignment (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Style/%E5%B1%9E%E6%80%A7/HorizontalAlignment%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: DGja_aM5
---

# HorizontalAlignment (属性)

返回或设置一个 XlHAlign 值，它代表指定对象的水平对齐方式。

## 说明

某些常量可能不可用，这取决于所选择或安装的语言支持（例如，美国英语）。

## 示例

```javascript
/*本示例为单元格 A1 赋值，并将单元格水平对齐方式设置为靠左对齐。*/
function test() {
    Range("A1").Value2 = "示例文本"
    Range("A1").Style.HorizontalAlignment = xlHAlignLeft
}
```

```javascript
/*本示例为单元格 D1 赋值，并显示D1单元格水平对齐方式是否为靠右对齐。*/
function test() {
    Range("D1").Value2 = "示例文本"
    console.log(Range("D1").Style.HorizontalAlignment == xlHAlignRight)
}
```
