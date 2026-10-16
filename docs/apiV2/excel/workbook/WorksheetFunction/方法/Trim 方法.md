---
title: "Trim 方法"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/WorksheetFunction/%E6%96%B9%E6%B3%95/Trim%20%E6%96%B9%E6%B3%95.html
---
# Trim (方法)

删除文本中的所有空格（除了单词之间的单个空格外）。对于从其他应用程序中获取带有不规则空格的文本，可以使用 TRIM。

## 说明

TRIM 函数设计用于从文本中剪裁掉 7 位 ASCII 空格字符（值 32）。在 Unicode 字符集中，有称为不间断空格字符的其他空格字符，其十进制值为 160。此字符通常在网页中作为 HTML 实体 使用。TRIM 函数本身并不会删除此不间断空格字符。

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| Arg1 | string | 必填 | Text - 要从中删除空格的文本。 |

## 返回值

String

## 示例

```javascript
/*本示例使用 Trim 方法删除文本中的所有空格，并分别赋值给A1和A2单元格*/
function test() {
    Range("A1").Value2 = WorksheetFunction.Trim("你好    欢迎您")
    Range("A2").Value2 = WorksheetFunction.Trim("How are you?")
}
```

```javascript
/*本示例为A2单元格赋值，使用 Trim 方法删除文本中的所有空格。*/
function test() {
    Range("A2").Value2 = "Hello    Wps"
    console.log(WorksheetFunction.Trim(Range("A2").Value2))
}
```
