---
title: "HorizontalAlignment (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Range/%E5%B1%9E%E6%80%A7/HorizontalAlignment%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: oNPBZGra
---

# HorizontalAlignment (属性)

返回或设置一个 **Variant** 值，它代表指定对象的水平对齐方式。

## 说明

此属性的值可设为以下常量之一：

| **xlCenter** |
| --- |
| **xlDistributed** |
| **xlJustify** |
| **xlLeft** |
| **xlRight** |

某些常量可能不可用，这取决于所选择或安装的语言支持（例如，美国英语）。

## 示例

```javascript
/*本示例将 A1 单元格的对齐方式设置为居中*/
function test() {
    Range("A1").HorizontalAlignment = xlCenter
}
```

```javascript
/*本示例将 B1 单元格的对齐方式设置为左对齐*/
function test() {
    Range("B1").HorizontalAlignment = xlLeft
}
```
