---
title: "MIrr 方法"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/WorksheetFunction/%E6%96%B9%E6%B3%95/MIrr%20%E6%96%B9%E6%B3%95.html
---
# MIrr (方法)

返回某一连续期间内现金流的修正内部收益率。函数 MIRR 同时考虑了投资的成本和现金再投资的收益率。

## 说明

- 参数 Values 中必须至少包含一个正值和一个负值，才能计算修正后的内部收益率，否则函数 MIRR 会返回错误值 #DIV/0!。
- 如果数组或引用参数包含文本、逻辑值或空单元格，则这些值将被忽略；但含有零值的单元格包括在内。
- 函数 MIRR 根据输入值的次序来解释现金流的次序。所以，务必按照实际的顺序输入支出和收入数额，并使用正确的正负号（现金流入用正值，现金流出用负值）。

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| Arg1 | any | 必填 | Values - 一个数组或对包含数字的单元格的引用。这些数值代表固定期间的一系列支出（负值）及收入（正值）。 |
| Arg2 | double | 必填 | Finance_rate - 现金流中使用的资金支付的利率。 |
| Arg3 | double | 必填 | Reinvest_rate - 将现金流再投资的收益率。 |

## 返回值

Double

## 示例

```javascript
/*本示例使用 MIrr 方法计算连续期间内现金流的修正内部收益率，并将结果分别赋值给D1和D2单元格。*/
function test() {
    let arr1 = [100, 2000, -453.2, -115, 126, -110.6]
    Range("D1").Value2 = WorksheetFunction.MIrr(arr1, 0.024, 0.0012)
    Range("D2").Value2 = WorksheetFunction.MIrr(arr1, 0.015, 0.016)
}
```

```javascript
/*本示例为A1等单元格分别赋值，使用 MIrr 方法计算连续期间内现金流的修正内部收益率。*/
function test() {
    Range("A1").Value2 = 104
    Range("A2").Value2 = 105
    Range("A3").Value2 = 106
    Range("A4").Value2 = -110
    Range("B1").Value2 = 0.12
    Range("B2").Value2 = 0.26
    console.log(WorksheetFunction.MIrr(Range("A1:A4"), Range("B1").Value2, Range("B2").Value2))
}
```
