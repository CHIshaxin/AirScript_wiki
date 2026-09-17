---
title: "Npv 方法"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/WorksheetFunction/%E6%96%B9%E6%B3%95/Npv%20%E6%96%B9%E6%B3%95.html
---
# Npv (方法)

通过使用贴现率以及一系列未来支出（负值）和收入（正值），计算一项投资的净现值。

## 说明

- Value1, value2, ... 在时间上必须具有相等间隔，并且都发生在期末。
- NPV 使用 value1, value2, ... 的顺序来解释现金流的顺序。所以务必保证支出和收入的数额按正确的顺序输入。
- 如果参数为数值、空白单元格、逻辑值或数字的文本表达式，则都会计算在内；如果参数是错误值或不能转化为数值的文本，则被忽略。
- 如果参数是一个数组或引用，则只计算其中的数字。数组或引用中的空白单元格、逻辑值、文本或错误值将被忽略。
- 函数 NPV 假定投资开始于 value1 现金流所在日期的前一期，并结束于列表中最后一笔现金流的当期。函数 NPV 依据未来的现金流来进行计算。如果第一笔现金流发生在第一个周期的期初，则第一笔现金必须添加到函数 NPV 的结果中，而不应包含在 values 参数中。有关详细信息，请参阅下面的示例。
- 函数 NPV 与函数 PV（现值）相似。PV 与 NPV 之间的主要差别在于：函数 PV 允许现金流在期初或期末开始。与可变的 NPV 的现金流数值不同，PV 的每一笔现金流在整个投资中必须是固定的。有关年金与财务函数的详细信息，请参阅函数 PV。
- 函数 NPV 与函数 IRR（内部收益率）也有关，函数 IRR 是使 NPV 等于零的比率：NPV(IRR(...), ...) = 0。

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| Arg1 | double | 必填 | Rate - 某一期间的贴现率。 |
| Arg2 | any | 必填 | Arg30 - Value1, value2, ... - 代表支出及收入的 1 到 29 个参数。 |

## 返回值

Double

## 示例

```javascript
/*本示例使用 Npv 方法计算一项投资的净现值，并分别赋值给B2和B3单元格。*/
function test() {
    Range("B2").Value2 = WorksheetFunction.Npv(0.235, 300.36, 600, -253, -68)
    Range("B3").Value2 = WorksheetFunction.Npv(0.36, 2000.36, -168)
}
```

```javascript
/*本示例为A1等单元格分别赋值，使用 Npv 方法计算一项投资的净现值。*/
function test() {
    Range("A1").Value2 = 0.125
    Range("A2").Value2 = 90000
    Range("A3").Value2 = -325.03
    console.log(WorksheetFunction.Npv(Range("A1").Value2, Range("A2").Value2, Range("A3").Value2))
}
```
