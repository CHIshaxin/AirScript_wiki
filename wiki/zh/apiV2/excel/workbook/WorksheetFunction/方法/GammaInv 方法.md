---
title: "GammaInv (方法)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/WorksheetFunction/%E6%96%B9%E6%B3%95/GammaInv%20%E6%96%B9%E6%B3%95.html
updated: 1789700916000
hash: DD93RG0N
---

# GammaInv (方法)

返回 γ 累积分布函数的反函数。如果 p = GAMMADIST(x,...)，则 GAMMAINV(p,...) = x。

## 说明

可以使用此函数研究可能出现偏态分布的变量。

- 如果任一参数为文本，则 GAMMAINV 将返回错误值 #VALUE!。
- 如果 probability < 0 或 probability > 1，则 GAMMAINV 将返回错误值 #NUM!。
- 如果 alpha ≤ 0 或 beta ≤ 0，则 GAMMAINV 将返回错误值 #NUM!。

如果已给定概率值，则 GAMMAINV 使用 GAMMADIST(x, alpha, beta, TRUE) = probability 求解数值 x。因此，GAMMAINV 的精度取决于 GAMMADIST 的精度。GAMMAINV 使用迭代搜索技术。如果搜索在 64 次迭代之后没有收敛，则该函数将返回错误值 #N/A。

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| Arg1 | double | 必填 | probability - 与 γ 分布相关联的概率。 |
| Arg2 | double | 必填 | alpha - 分布参数。 |
| Arg3 | double | 必填 | beta - 分布参数。如果 beta = 1，则 GAMMAINV 将返回标准 γ 分布。 |

## 返回值

Double

## 示例

```javascript
/*本示例使用 GammaInv 方法计算γ累积分布函数的反函数，并分别赋值给D1和D2单元格。*/
function test() {
    Range("D1").Value2 = WorksheetFunction.GammaInv(0, 0.6, 0.4)
    Range("D2").Value2 = WorksheetFunction.GammaInv(0.2, 6, 1)
}
```

```javascript
/*本示例为B1等单元格分别赋值，使用 GammaInv 方法计算γ累积分布函数的反函数。*/
function test() {
    Range("B1").Value2 = 0.99
    Range("B2").Value2 = 0.5
    Range("B3").Value2 = 6
    console.log(WorksheetFunction.GammaInv(Range("B1").Value2, Range("B2").Value2, Range("B3").Value2))
}
```
