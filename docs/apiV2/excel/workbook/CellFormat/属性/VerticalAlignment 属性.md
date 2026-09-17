---
title: "VerticalAlignment 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/CellFormat/%E5%B1%9E%E6%80%A7/VerticalAlignment%20%E5%B1%9E%E6%80%A7.html
---
# VerticalAlignment (属性)

返回或设置一个 **Variant** 值，它代表指定对象的垂直对齐方式。

## 说明

此属性的值可以设置为 **XlVAlign** 常量之一。

## 示例

```javascript
/*此示例显示查找条件中的垂直对齐方式是否为 xlVAlignDistributed 。*/
function test() {
    console.log(Application.FindFormat.VerticalAlignment == xlVAlignDistributed)
}
```

```javascript
/*此示例将替换条件设置为向下垂直对齐。*/
function test() {
    Application.ReplaceFormat.VerticalAlignment = xlVAlignBottom
}
```
