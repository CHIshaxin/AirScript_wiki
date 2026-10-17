---
title: "Min (方法)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/WorksheetFunction/%E6%96%B9%E6%B3%95/Min%20%E6%96%B9%E6%B3%95.html
updated: 1789700916000
hash: CZ8_NuFW
---

# Min (方法)

返回一组值中的最小值。

## 说明

- 参数可以是数字，也可以是包含数字的名称、数组或引用。
- 直接键入参数列表的数字的逻辑值和文本表示也包括在内。
- 如果参数为数组或引用，则只使用其中的数值。数组或引用中的空白单元格、逻辑值或文本将被忽略。
- 如果参数不包含任何数字，则 MIN 返回 0。
- 如果参数为错误值或不能转换为数字的文本，则将导致错误。
- 如果想在引用中将数字的逻辑值和文本表示作为计算的一部分包括在内，请使用 MINA 函数。

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| Arg1 | any | 必填 | Arg30 - Number1, number2, ... - 要从中查找最小值的 1 到 30 个数字。 |

## 返回值

Double

## 示例

```javascript
/*本示例使用 Min 方法计算一组值中的最小值，并将结果分别赋值给A1和A2单元格。*/
function test() {
    let arr1 = [1, 2, 3, 4]
    Range("A1").Value2 = WorksheetFunction.Min(arr1)
    Range("A2").Value2 = WorksheetFunction.Min(10, 20, -4, 2.6, 5)
}
```

```javascript
/*本示例为A1等单元格分别赋值，使用 Min 方法计算区域A1:A4中的最小值。*/
function test() {
    Range("A1").Value2 = 10.6
    Range("A2").Value2 = 10.65
    Range("A3").Value2 = 10.23
    Range("A4").Value2 = 10.2
    console.log(WorksheetFunction.Min(Range("A1:A4")))
}
```
