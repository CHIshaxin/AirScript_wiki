---
title: "Or (方法)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/WorksheetFunction/%E6%96%B9%E6%B3%95/Or%20%E6%96%B9%E6%B3%95.html
updated: 1789700916000
hash: k7VpJode
---

# Or (方法)

在其参数组中，如果任一参数的逻辑值为 TRUE，即返回 TRUE；如果所有参数的逻辑值为 FALSE，即返回 FALSE。

## 说明

- 参数必须能计算为逻辑值，如 TRUE 或 FALSE，或者为包含逻辑值的数组 （数组：用于建立可生成多个结果或可对在行和列中排列的一组参数进行运算的单个公式。数组区域共用一个公式；数组常量是用作参数的一组常量。）或引用。
- 如果数组或引用参数中包含文本或空单元格，则这些值将被忽略。
- 如果指定的区域中不包含任何逻辑值，则 OR 返回错误值 #VALUE!。
- 可以使用 OR 数组公式来检验数组中是否包含特定的数值。若要输入数组公式，请按 Ctrl+Shift+Enter。

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
| --- | --- | --- | --- |

## 返回值

Boolean

## 示例

```javascript
/*本示例使用 Or 方法判断参数的逻辑值，并分别赋值给C1和C2单元格。*/
function test() {
    Range("C1").Value2 = WorksheetFunction.Or(true, false, false, false)
    Range("C2").Value2 = WorksheetFunction.Or(false, false)
}
```

```javascript
/*本示例为A1等单元格分别赋值，使用 Or 方法判断如下相应条件的逻辑值。*/
function test() {
    Range("A1").Value2 = 4
    Range("A2").Value2 = 3
    Range("A3").Value2 = 0
    Range("A4").Value2 = 10
    console.log(WorksheetFunction.Or(Range("A1").Value2 > 1, Range("A2").Value2 > 1, Range("A3").Value2 < 1, Range("A4").Value2 == 10))
}
```
