---
title: "GammaLn 方法"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/WorksheetFunction/%E6%96%B9%E6%B3%95/GammaLn%20%E6%96%B9%E6%B3%95.html
---
# GammaLn (方法)

返回 γ 函数的自然对数 Γ(x)。

## 说明

- 如果 x 是非数值型，则 GAMMALN 将返回错误值 #VALUE!。
- 如果 x ≤ 0，则 GAMMALN 将返回错误值 #NUM!。
- 数字 e 的 GAMMALN(i) 次幂的返回值与 (i - 1)! 的结果相同（其中 i 为整数）。

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| Arg1 | double | 必填 | X - 要计算其 GAMMALN 值的数值。 |

## 返回值

Double

此函数已被一个或多个新函数取代，这些新函数可以提供更高的准确度，而且它们的名称可以更好地反映出其用途。仍然提供此函数是为了保持与 ET 早期版本的兼容性。但是，如果不需要后向兼容性，则应考虑从现在开始使用新函数，因为它们可以更加准确地描述其功能。

有关新函数的详细信息，请参阅 **GammaLn_Precise** 方法。

## 示例

```javascript
/*本示例使用 GammaLn 方法计算 γ 函数的自然对数 Γ(x)，并分别赋值给B1和B2单元格。*/
function test() {
    Range("B1").Value2 = WorksheetFunction.GammaLn(2)
    Range("B2").Value2 = WorksheetFunction.GammaLn(Math.E)
}
```

```javascript
/*本示例为C1单元格赋值，使用 GammaLn 方法计算C1单元格的 γ 函数的自然对数 Γ(x)。*/
function test() {
    Range("C1").Value2 = 3.3
    console.log(WorksheetFunction.GammaLn(Range("C1").Value2))
}
```
