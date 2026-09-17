---
title: "Pmt 方法"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/WorksheetFunction/%E6%96%B9%E6%B3%95/Pmt%20%E6%96%B9%E6%B3%95.html
---
# Pmt (方法)

基于固定利率及等额分期付款方式，计算贷款的每期付款额。

## 说明

有关函数 PMT 中参数的详细说明，请参阅函数 PV。

| Type 值 | 付款时间 |
| --- | --- |
| 0 或省略 | 期末 |
| 1 | 期初 |

- PMT 返回的支付款项包括本金和利息，但不包括税款、保留支付或某些与贷款有关的费用。
- 应确保所指定的 rate 和 nper 单位的一致性。例如，同样是四年期年利率为 12% 的贷款，如果按月支付，rate 应为 12%/12，nper 应为 4\*12；如果按年支付，rate 应为 12%，nper 为 4。

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| Arg1 | double | 必填 | Rate - 贷款的利率。 |
| Arg2 | double | 必填 | Nper - 该项贷款的付款总次数。 |
| Arg3 | double | 必填 | Pv - 现值，或一系列未来付款的当前值的累积和，也称为本金。 |
| Arg4 | any | 可选 | Fv - 未来值，或在最后一次付款后希望得到的现金余额。如果省略 fv，则假定其值为 0（零），即一笔贷款的未来值为 0。 |
| Arg5 | any | 可选 | Type - 数字 0（零）或 1，用以指定各期的付款时间是在期初还是期末。 |

## 返回值

Double

## 示例

```javascript
/*本示例使用 Pmt 方法分别计算贷款5次和8次的每期付款额，并分别赋值给A1和A2单元格。*/
function test() {
    Range("A1").Value2 = WorksheetFunction.Pmt(0.0013, 5, 122, 0)
    Range("A2").Value2 = WorksheetFunction.Pmt(0.0053, 8, 242, 21, 1)
}
```

```javascript
/*本示例为D1等单元格分别赋值，使用 Pmt 方法计算贷款的每期付款额。*/
function test() {
    Range("D1").Value2 = 0.0011
    Range("D2").Value2 = 3
    Range("D3").Value2 = 480
    Range("D4").Value2 = 3200
    console.log(WorksheetFunction.Ppmt(Range("D1").Value2, Range("D2").Value2, Range("D3").Value2, Range("D4").Value2))
}
```
