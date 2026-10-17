---
title: "Xnpv (方法)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/WorksheetFunction/%E6%96%B9%E6%B3%95/Xnpv%20%E6%96%B9%E6%B3%95.html
updated: 1789700916000
hash: Cys8lH6h
---

# Xnpv (方法)

返回一组现金流的净现值，这些现金流不一定定期发生。可读/写 **Double** 类型。

## 说明

若要计算一组定期现金流的净现值，请使用 **Npv** 方法。

**Xnpv** 方法不提供与相应的 **XNPV** 函数 (=XNPV(**rate**, values, dates)) 所需的 rate 参数对应的参数。若要在 VBA 代码中避开这一限制，可通过使用下面所示的 **Evaluate** 方法来调用 **XNPV** 函数，而不使用 **Xpnv** 方法。

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| Arg1 | any | 必填 | 与 dates 中的付款计划相对应的一系列现金流。首期付款是可选的，并与投资开始时的成本或付款相对应。 |
| Arg2 | any | 必填 | 与现金流付款相对应的付款日期表。第一个付款日期代表付款计划的开始日期。所有其他日期必须迟于该日期，但可按任何顺序排列。 |

## 返回值

Double

## 示例

以下示例返回上面的成本和收益下的投资净现值。现金流的贴现率为 9%。（2086.6476 或 2086.65）。

let npv = Application.Evaluate("=XNPV(.09,A2:A6,B2:B6)")

```javascript
/*本示例演示用 Xnpv 计算一组现金流的净现值，并将结果赋值到B5单元格*/
function test() {
    let dts = [100.1, 1000.2, 2000.52]
    Range("D3").Value2 = "2021/09/10"
    Range("D4").Value2 = "2023/09/30"
    Range("D5").Value2 = "2024/11/14"
    Range("B5").Value2 = WorksheetFunction.Xnpv(dts, Range("D3:D5"))
}
```

```javascript
/*本示例演示为C1等单元格分别赋值，用 Xnpv 计算一组现金流的净现值。*/
function test() {
    Range("C1").Value2 = -1000
    Range("C2").Value2 = 2000
    Range("C3").Value2 = 3000
    Range("D1").Value2 = "2021/09/10"
    Range("D2").Value2 = "2023/09/30"
    Range("D3").Value2 = "2024/11/14"
    console.log(WorksheetFunction.Xnpv(Range("C1:C3"), Range("D1:D3")))
}
```
