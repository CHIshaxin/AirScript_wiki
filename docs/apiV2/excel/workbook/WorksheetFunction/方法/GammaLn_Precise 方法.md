---
title: "GammaLn_Precise 方法"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/WorksheetFunction/%E6%96%B9%E6%B3%95/GammaLn_Precise%20%E6%96%B9%E6%B3%95.html
---
# GammaLn_Precise (方法)

返回 γ 函数的自然对数 Γ(x)。

## 说明

- 如果 x 为非数值，**GammaLn_Precise** 方法将返回错误。
- 如果 x ≤ 0，**GammaLn_Precise** 方法将返回错误。
- 数字 e 的 GAMMALN(i)（i 为整数）次幂的返回值与 (i - 1)! 的结果相同。

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| Arg1 | double | 必填 | X - 要计算其 GAMMALN 的值。 |

## 返回值

Double

## 示例

```javascript
/*本示例使用 GammaLn_Precise 方法计算 γ 函数的自然对数 Γ(x)，并分别赋值给A1和A2单元格。*/
function test() {
    Range("A1").Value2 = WorksheetFunction.GammaLn_Precise(1.2)
    Range("A2").Value2 = WorksheetFunction.GammaLn_Precise(Math.E)
}
```

```javascript
/*本示例为D1单元格赋值，使用 GammaLn_Precise 方法计算D1单元格的 γ 函数的自然对数 Γ(x)。*/
function test() {
    Range("D1").Value2 = 3.3
    console.log(WorksheetFunction.GammaLn_Precise(Range("D1").Value2))
}
```
