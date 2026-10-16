---
title: "Xirr 方法"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/WorksheetFunction/%E6%96%B9%E6%B3%95/Xirr%20%E6%96%B9%E6%B3%95.html
---
# Xirr (方法)

返回一组现金流的内部收益率，这些现金流不一定定期发生。若要计算一组定期现金流的内部收益率，请使用函数 IRR。

## 说明

- ET 以序数形式存储日期以使其可用于计算。默认情况下，1900 年 1 月 1 日的序数是 1；2008 年 1 月 1 日的序数是 39448，因为该日期距 1900 年 1 月 1 日有 39,448 天。
- dates 中的数字将被截尾取整。
- XIRR 要求至少有一个正现金流和一个负现金流，否则 XIRR 将返回错误值 #NUM!。
- 如果 dates 中的任一数字不是有效日期，XIRR 将返回错误值 #VALUE。
- 如果 dates 中的任一数字早于开始日期，XIRR 将返回错误值 #NUM!。
- 如果 values 和 dates 所含的值的个数不同，XIRR 将返回错误值 #NUM!。
- 多数情况下，不需要为 XIRR 的计算提供 guess 值。guess 的值在省略时假定为 0.1 (10％)。
- XIRR 与净现值函数 XNPV 密切相关。XIRR 计算的收益率是 XNPV = 0 时的利率。
- ET 使用迭代技术计算 XIRR。XIRR 使用不断修正的收益率（从 guess 开始）循环执行计算，直到结果的精度达到 0.000001% 以内为止。如果 XIRR 执行了 100 次循环后还是找不到结果，则将返回错误值 #NUM!。

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| Arg1 | any | 必填 | Values - 与 dates 中的付款计划对应的一系列现金流。第一个付款是可选的，与投资开始时的成本或付款对应。第一个值如果是成本或付款，就必须是负值。所有后续付款都基于 365 天/年贴现。这一系列值中必须至少包含一个正值和一个负值。 |
| Arg2 | any | 必填 | Dates - 与现金流付款对应的付款日期计划。第一个付款日期表示付款计划的开始。所有其他日期都必须晚于此日期，但相互顺序是任意的。日期应使用 DATE 函数输入，或者作为其他公式或函数的结果输入。例如，使用 DATE(2008,5,23) 输入 2008 年 5 月 23 日。如果日期以文本形式输入，将会出现问题。 |
| Arg3 | any | 可选 | Guess - 一个据您推测是接近 XIRR 结果的数字。 |

## 返回值

Double

## 示例

```javascript
/*本示例演示计算一组现金流的内部收益率，并将结果赋值到A5单元格*/
function test() {
    let dts = [-1000, 2000, 2400, 2880]
    let dats = ["2022/09/10", "2022/10/10", "2023/11/10", "2023/12/10"]
    Range("A5").Value2 = WorksheetFunction.Xirr(dts, dats, 11)
}
```

```javascript
/*本示例演示为A1等单元格分别赋值，用 Xirr 计算一组现金流的内部收益率，这些现金流不一定定期发生。*/
function test() {
    Range("A1").Value2 = -1000
    Range("A2").Value2 = 2000
    Range("A3").Value2 = 3000
    Range("B1").Value2 = "2021/09/10"
    Range("B2").Value2 = "2023/09/30"
    Range("B3").Value2 = "2024/11/14"
    console.log(WorksheetFunction.Xirr(Range("A1:A3"), Range("B1:B3")))
}
```
