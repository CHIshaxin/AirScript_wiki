---
title: "Rate (方法)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/WorksheetFunction/%E6%96%B9%E6%B3%95/Rate%20%E6%96%B9%E6%B3%95.html
updated: 1789700916000
hash: UUTM7DcU
---

# Rate (方法)

返回年金的各期利率。函数 RATE 通过迭代法计算得出，并且可能无解或有多个解。如果在进行 20 次迭代计算后，函数 RATE 的相邻两次结果没有收敛于 0.0000001，则函数 RATE 将返回错误值 #NUM!。

## 说明

有关参数 nper、pmt、pv、fv 和 type 的完整说明，请参阅 PV。

| Type 值 | 付款时间 |
| --- | --- |
| 0 或省略 | 期末 |
| 1 | 期初 |

- 如果省略 guess，则假定其值为 10%。
- 如果 RATE 不收敛，则请尝试使用不同的 guess 值。当 guess 介于 0 和 1 之间时，RATE 通常会收敛。

应确保所指定的 guess 和 nper 单位的一致性。例如，同样是四年期年利率为 12% 的贷款，如果按月支付，guess 应为 12%/12，nper 应为 4*12；如果按年支付，guess 应为 12%，nper 应为 4。

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| Arg1 | double | 必填 | Nper - 年金中的总付款期数。 |
| Arg2 | double | 必填 | Pmt - 各期付款额，其数值在年金的整个时间期限内保持不变。通常 pmt 包括本金和利息，但不包括其他费用或税金。如果省略 pmt，则必须包含 fv 参数。 |
| Arg3 | double | 必填 | Pv - 现值，即一系列未来付款当前所值的总额。 |
| Arg4 | any | 可选 | Fv - 未来值，或在最后一次付款后希望得到的现金余额。如果省略 fv，则假定其值为 0（例如，一笔贷款的未来值为 0）。 |
| Arg5 | any | 可选 | Type - 数字 0 或 1，用以指示付款时间。 |
| Arg6 | any | 可选 | Guess - 猜测的利率。 |

## 返回值

Double

## 示例

```javascript
/*本示例使用 Rate 方法分别计算金额10000和金额15000在36期和60期每期年金的利率，并赋值给B1和B2单元格。*/
function test() {
    Range("B1").Value2 = WorksheetFunction.Rate(36, -300, 10000)
    Range("B2").Value2 = WorksheetFunction.Rate(60, -400, 15000, 10, 1, 0.2)
}
```

```javascript
/*本示例为A1等单元格分别赋值，使用 Rate 方法计算以A1单元格的值为期数，A2单元格的值为付款额，A3单元格的值为现值的每期年金利率。*/
function test() {
    Range("A1").Value2 = 72
    Range("A2").Value2 = -350
    Range("A3").Value2 = 20000
    console.log(WorksheetFunction.Rate(Range("A1").Value2, Range("A2").Value2, Range("A3").Value2))
}
```
