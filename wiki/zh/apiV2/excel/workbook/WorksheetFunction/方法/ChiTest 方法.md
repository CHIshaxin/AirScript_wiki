---
title: "ChiTest (方法)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/WorksheetFunction/%E6%96%B9%E6%B3%95/ChiTest%20%E6%96%B9%E6%B3%95.html
updated: 1789700916000
hash: B_J3iV1s
---

# ChiTest (方法)

返回独立性检验值。

## 说明

ChiTest 返回 χ2 分布的统计值及相应的自由度。可以使用 χ2 检验确定假设结果是否被实验所证实。

- 如果 actual_range 和 expected_range 数据点的个数不同，则 ChiTest 将返回错误值 #N/A。
- χ2 的低值是独立指示器。从公式中可看出，χ2 总是正数或 0，并且对于每个 ij，仅在 Aij = Eij 时为 0。
- ChiTest 返回在独立的假设条件下意外获得特定情况的概率，即 χ2 统计值至少和由上面的公式计算出的值一样大的情况。在计算此概率时，ChiTest 会使用具有相应自由度 df 的 χ2 分布。如果 r > 1 且 c > 1，则 df = (r - 1)(c - 1)。如果 r = 1 且 c > 1，则 df = c - 1。如果 r > 1 且 c = 1，则 df = r - 1。不允许出现 r = c= 1，此时将生成一个错误。
- 当 Eij 的值不太小时，使用 ChiTest 最合适。一些统计学家建议每个 Eij 应该大于等于 5。

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| Arg1 | any | 必填 | 包含观察值的数据区域，将根据期望值检验这些观察值。 |
| Arg2 | any | 必填 | 包含行汇总和列汇总的乘积与总计值之比率的数据区域。 |

## 返回值

Double

## 示例

```javascript
/*本示例为A2等单元格分别赋值，并使用ChiTest方法计算独立性检验值并显示结果。*/
function test() {
    Range("A2").Value2 = 58
    Range("A3").Value2 = 11
    Range("A4").Value2 = 45.35
    Range("A5").Value2 = 17.56
    Range("B2").Value2 = 35
    Range("B3").Value2 = 25
    Range("B4").Value2 = 47.65
    Range("B5").Value2 = 18.44
    let chitest1 = Application.WorksheetFunction.ChiTest(Range("A2:B3"), Range("A4:B5"))
    console.log(chitest1)
}
```

```javascript
/*本示例为C2等单元格分别赋值，并使用ChiTest方法计算独立性检验值，并将结果赋值于F2单元格。*/
function test() {
    Range("C2").Value2 = 47
    Range("C3").Value2 = 36
    Range("C4").Value2 = 20
    Range("D2").Value2 = 45.33
    Range("D3").Value2 = 12.5
    Range("D4").Value2 = 26.54
    Range("F2").Value2 = Application.WorksheetFunction.ChiTest(Range("C2:C4"), Range("D2:D4"))
}
```
