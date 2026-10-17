---
title: "Proper (方法)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/WorksheetFunction/%E6%96%B9%E6%B3%95/Proper%20%E6%96%B9%E6%B3%95.html
updated: 1789700916000
hash: CUeUK3Zi
---

# Proper (方法)

将文本字符串的首字母及任何非字母字符之后的首字母转换成大写。将其余的字母转换成小写。

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| Arg1 | string | 必填 | Text - 用引号括起来的文本、返回文本的公式或是对包含要进行部分大写转换的文本的单元格引用。 |

## 返回值

String

## 示例

```javascript
/*本示例使用 Proper 方法将字符串WorksheetFunction和HELLO的首字母及任何非字母字符之后的首字母转换成大写，并分别赋值给B1和B2单元格。*/
function test() {
    Range("B1").Value2 = WorksheetFunction.Proper("WorksheetFunction")
    Range("B2").Value2 = WorksheetFunction.Proper("HELLO")
}
```

```javascript
/*本示例为A1单元格赋值，使用 Proper 方法将A1单元格中字符串的首字母及任何非字母字符之后的首字母转换成大写。*/
function test() {
    Range("A1").Value2 = "worksheet"
    console.log(WorksheetFunction.Proper(Range("A1").Value2))
}
```
