---
title: "Roman (方法)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/WorksheetFunction/%E6%96%B9%E6%B3%95/Roman%20%E6%96%B9%E6%B3%95.html
updated: 1789700916000
hash: CyqEQiDt
---

# Roman (方法)

将阿拉伯数字转换为文本形式的罗马数字。

## 说明

| Form | 类型 |
| --- | --- |
| 0 或省略 | 古典型。 |
| 1 | 更简明。请参阅下面的示例。 |
| 2 | 更简明。请参阅下面的示例。 |
| 3 | 更简明。请参阅下面的示例。 |
| 4 | 简化型。 |
| TRUE | 古典型。 |
| FALSE | 简化型。 |

- 如果 number 为负数，则将返回错误值 #VALUE!。
- 如果 number 大于 3999，则将返回错误值 #VALUE!。

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| Arg1 | double | 必填 | Number - 要转换的阿拉伯数字。 |
| Arg2 | any | 可选 | Form - 一个数字，指定所需的罗马数字类型。罗马数字的样式范围可以从古典型到简化，随着 form 值的增加将趋于更简明。 |

## 返回值

String

## 示例

```javascript
/*本示例使用 Roman 方法分别将阿拉伯数字11和45转换为文本形式的罗马数字，并分别赋值给A1和A2单元格。*/
function test() {
    Range("A1").Value2 = WorksheetFunction.Roman(11)
    Range("A2").Value2 = WorksheetFunction.Roman(45, 4)
}
```

```javascript
/*本示例为C1等单元格分别赋值，使用 Roman 方法将C1单元格中阿拉伯数字转换为文本形式的罗马数字。*/
function test() {
    Range("C1").Value2 = 25
    Range("D1").Value2 = 2
    console.log(WorksheetFunction.Roman(Range("C1").Value2, Range("D1").Value2))
}
```
