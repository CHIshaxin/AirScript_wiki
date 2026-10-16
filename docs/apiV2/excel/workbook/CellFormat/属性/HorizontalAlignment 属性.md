---
title: "HorizontalAlignment 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/CellFormat/%E5%B1%9E%E6%80%A7/HorizontalAlignment%20%E5%B1%9E%E6%80%A7.html
---
# HorizontalAlignment (属性)

返回或设置一个 **Variant** 值，它代表指定对象的水平对齐方式。

## 说明

此属性的值可以设置为 **XlHAlign** 常量之一。

某些常量可能不可用，这取决于所选择或安装的语言支持（例如，美国英语）。

## 示例

```javascript
/*此示例设置搜索条件中对象的水平对其方式为居中。*/
function test() {
    Application.FindFormat.HorizontalAlignment = xlHAlignCenter
}
```

```javascript
/*此示例显示替换条件中水平对齐方式是否为两端对齐。*/
function test() {
    console.log(Application.ReplaceFormat.HorizontalAlignment == xlHAlignJustify)
}
```
