---
title: "AddIndent 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Style/%E5%B1%9E%E6%80%A7/AddIndent%20%E5%B1%9E%E6%80%A7.html
---
# AddIndent (属性)

返回或设置一个 **Boolean** 值，它指明当单元格中文本的对齐方式为水平或垂直等距分布时，文本是否为自动缩进。

## 说明

如果将此属性的值设为 **True**，那么在单元格中文本的对齐方式设为水平或垂直等距分布时，将自动缩进文本。

要将文本对齐方式设为等距分布，可在 **Orientation** 属性的值为 **xlVertical** 时，将 **VerticalAlignment** 属性设为 **xlVAlignDistributed**；在 **Orientation** 属性的值为 **xlHorizontal** 时，将 **HorizontalAlignment** 属性设为 **xlHAlignDistributed**。

## 示例

```javascript
/*本示例为A1单元格添加文本，并显示A1单元格文本是否自动缩进。*/
function test() {
    Range("A1").Value2 = "示例文本1"
    console.log(Range("A1").Style.AddIndent)
}
```

```javascript
/*本示例将B1单元格的文本设置为自动缩进。*/
function test() {
    Range("B1").Style.AddIndent = true
}
```
