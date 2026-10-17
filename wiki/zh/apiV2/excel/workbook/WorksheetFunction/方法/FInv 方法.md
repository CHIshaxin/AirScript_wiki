---
title: "FInv (方法)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/WorksheetFunction/%E6%96%B9%E6%B3%95/FInv%20%E6%96%B9%E6%B3%95.html
updated: 1789700916000
hash: D-C7SeJS
---

# FInv (方法)

返回 F 概率分布的反函数。如果 p = FDIST(x,...)，则 FINV(p,...) = x。

## 说明

在 F 检验中，可以使用 F 分布比较两个数据集的变化程度。例如，可以分析美国、加拿大的收入分布，判断这两个国家的收入变化程度是否相似。

- 如果任一参数为非数值型，则 FINV 将返回错误值 #VALUE!。
- 如果 probability < 0 或 probability > 1，则 FINV 将返回错误值 #NUM!。
- 如果 degrees_freedom1 或 degrees_freedom2 不是整数，则将被截尾取整。
- 如果 degrees_freedom1 < 1 或 degrees_freedom1 ≥ 10^10，则 FINV 将返回错误值 #NUM!。
- 如果 degrees_freedom2 < 1 或 degrees_freedom2 ≥ 10^10，则 FINV 将返回错误值 #NUM!。

FINV 可以用于返回 F 分布的临界值。例如，ANOVA 计算的结果通常包含 F 统计值、F 概率和显著水平为 0.05 的 F 临界值等数据。若要返回 F 的临界值，请将显著水平参数用作为 FINV 的 probability 参数。

如果已给定概率值，则 FINV 使用 FDIST(x, degrees_freedom1, degrees_freedom2) = probability 求解数值 x。因此，FINV 的精度取决于 FDIST 的精度。FINV 使用迭代搜索技术。如果搜索在 64 次迭代之后没有收敛，则该函数将返回错误值 #N/A。

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| Arg1 | double | 必填 | probability - 与 F 累积分布相关联的概率。 |
| Arg2 | double | 必填 | degrees_freedom1 - 分子的自由度。 |
| Arg3 | double | 必填 | degrees_freedom2 - 分母的自由度。 |

## 返回值

Double

## 示例

```javascript
/*本示例使用 FInv 方法计算 F 概率分布的反函数结果，并分别赋值给B1等单元格。*/
function test() {
    Range("B1").Value2 = WorksheetFunction.FInv(0.05, 10, 15)
    Range("B2").Value2 = WorksheetFunction.FInv(1, 5, 20)
}
```

```javascript
/*本示例分别赋值A1等单元格，使用 FInv 方法计算 F 概率分布的反函数结果。*/
function test() {
    Range("A1").Value2 = 0.99
    Range("A2").Value2 = 100
    Range("A3").Value2 = 999
    console.log(WorksheetFunction.FInv(Range("A1").Value2, Range("A2").Value2, Range("A3").Value2))
}
```
