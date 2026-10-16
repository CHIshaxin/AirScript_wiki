---
title: "Max 方法"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/WorksheetFunction/%E6%96%B9%E6%B3%95/Max%20%E6%96%B9%E6%B3%95.html
---
# Max (方法)

返回一组值中的最大值。

## 说明

- 参数可以是数字，也可以是包含数字的名称、数组或引用。
- 直接键入参数列表的数字的逻辑值和文本表示也包括在内。
- 如果参数为数组或引用，则只使用其中的数值。数组或引用中的空白单元格、逻辑值或文本将被忽略。
- 如果参数不包含任何数字，则 MAX 返回 0（零）。
- 如果参数为错误值或不能转换为数字的文本，则将导致错误。
- 如果想在引用中将数字的逻辑值和文本表示作为计算的一部分包括在内，请使用 MAXA 函数。

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| Arg1 | any | 必填 | Arg30 - Number1, number2, ... - 要从中查找最大值的 1 到 30 个数字。 |

## 返回值

Double

## 示例

```javascript
/*本示例使用 Max 方法计算一组值中的最大值，并分别赋值给D1和D2单元格。*/
function test() {
    let arr1 = [1, 2, 3, 4, 5, 6]
    Range("D1").Value2 = WorksheetFunction.Max(arr1)
    Range("D2").Value2 = WorksheetFunction.Max(50, 30, 54, 6)
}
```

```javascript
/*本示例为B2等单元格分别赋值，使用 Max 方法计算区域B2:B5值中的最大值。*/
function test() {
    Range("B2").Value2 = 20.6
    Range("B3").Value2 = 4
    Range("B4").Value2 = 20
    Range("B5").Value2 = 2
    console.log(WorksheetFunction.Max(Range("B2:B5")))
}
```
