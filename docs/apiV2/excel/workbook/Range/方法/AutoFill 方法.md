---
title: "AutoFill 方法"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Range/%E6%96%B9%E6%B3%95/AutoFill%20%E6%96%B9%E6%B3%95.html
---
# AutoFill (方法)

对指定区域中的单元格执行自动填充。

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| Destination | Range | 必填 | 要填充的单元格。目标区域必须包括源区域。 |
| Type | Default<XlAutoFillType, 0> | 可选 | 指定填充类型。 |

## 返回值

Variant

## 示例

```javascript
/*本示例以工作表 Sheet1 上的单元格区域 A1:A2 为基础，对单元格区域 A1:A20 进行自动填充。本示例最初，会在单元格 A1 中输入 1，在单元格 A2 中输入 2。*/
function test() {
    Range("A1").Value2 = 1
    Range("A2").Value2 = 2
    let sourceRange = Range("A1:A2")
    let fillRange = Range("A1:A20")
    sourceRange.AutoFill(fillRange)
}
```

```javascript
/*本示例将 A1 单元格的文本加粗，并将格式复制到 A1:A5 单元格区域中*/
function test() {
    Range("A1").Font.Bold = true
    Range("A1").AutoFill(Range("A1:A5"), xlFillFormats)
}
```
