---
title: "NPer (方法)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/WorksheetFunction/%E6%96%B9%E6%B3%95/NPer%20%E6%96%B9%E6%B3%95.html
updated: 1789700916000
hash: AR_JkY4P
---

# NPer (方法)

基于固定利率及等额分期付款方式，返回某项投资的总期数。

## 说明

| Type 值 | 付款时间 |
| --- | --- |
| 0 或省略 | 期末 |
| 1 | 期初 |

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| Arg1 | double | 必填 | Rate - 各期利率。 |
| Arg2 | double | 必填 | Pmt - 各期所应支付的金额，其数值在整个年金期间保持不变。通常，pmt 包括本金和利息，但不包括其他的费用或税款。 |
| Arg3 | double | 必填 | Pv - 现值，或一系列未来付款当前所值的累积和。 |
| Arg4 | any | 可选 | Fv - 未来值，或在最后一次付款后希望得到的现金余额。如果省略 fv，则假定其值为 0（例如，一笔贷款的未来值为 0）。 |
| Arg5 | any | 可选 | Type - 数字 0 或 1，用以指示付款时间。 |

## 返回值

Double

## 示例

```javascript
/*本示例使用 NPer 方法基于固定利率及等额分期付款方式，分别计算投资金额6000和9004的总期数，并分别赋值给D1和D2单元格。*/
function test() {
    Range("D1").Value2 = WorksheetFunction.NPer(0.32, 6000, 10000, 0, 0)
    Range("D2").Value2 = WorksheetFunction.NPer(0.0015, 9004, 1, 2, 1)
}
```

```javascript
/*本示例为C1等单元格分别赋值，使用 NPer 方法基于固定利率及等额分期付款方式，计算C2单元格中投资的总期数。*/
function test() {
    Range("C1").Value2 = 0.321
    Range("C2").Value2 = 1326.3
    Range("C3").Value2 = 6000
    Range("C4").Value2 = 25
    Range("C5").Value2 = 0
    console.log(WorksheetFunction.NPer(Range("C1").Value2, Range("C2").Value2, Range("C3").Value2, Range("C4").Value2, Range("C5").Value2))
}
```
