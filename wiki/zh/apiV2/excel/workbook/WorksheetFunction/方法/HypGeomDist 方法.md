---
title: "HypGeomDist (方法)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/WorksheetFunction/%E6%96%B9%E6%B3%95/HypGeomDist%20%E6%96%B9%E6%B3%95.html
updated: 1789700916000
hash: C2npo8_T
---

# HypGeomDist (方法)

返回超几何分布。在给定样本容量、样本总体成功次数和样本总体容量时，HYPGEOMDIST 返回样本取得给定成功次数的概率。使用 HYPGEOMDIST 可解决有限总体的问题，其中每个观察值或为成功或为失败，且其中给定样本容量的每一个子集有相等的发生概率。

## 说明

- 所有参数都将被截尾取整。
- 如果任一参数为非数值型，则 HYPGEOMDIST 将返回错误值 #VALUE!。
- 如果 sample_s < 0 或 sample_s 大于 number_sample 或 population_s 中的较小值，则 HYPGEOMDIST 将返回错误值 #NUM!。
- 如果 sample_s 小于 0 或 (number_sample - number_population + population_s) 中的较大值，则 HYPGEOMDIST 将返回错误值 #NUM!。
- 如果 number_sample ≤ 0 或 number_sample > number_population，则 HYPGEOMDIST 将返回错误值 #NUM!。
- 如果 population_s ≤ 0 或 population_s > number_population，则 HYPGEOMDIST 将返回错误值 #NUM!。
- 如果 number_population ≤ 0，则 HYPGEOMDIST 将返回错误值 #NUM!。 HYPGEOMDIST 用于在有限总体中进行不退回抽样的概率计算。

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| Arg1 | double | 必填 | sample_s - 样本中成功的次数。 |
| Arg2 | double | 必填 | number_sample - 样本的大小。 |
| Arg3 | double | 必填 | population_s - 总体中成功的次数。 |
| Arg4 | double | 必填 | number_population - 总体大小。 |

## 返回值

Double

## 示例

```javascript
/*本示例使用 HypGeomDist 方法计算超几何分布，并分别赋值给A1和A2单元格。*/
function test() {
    Range("A1").Value2 = WorksheetFunction.HypGeomDist(0, 5, 1, 7)
    Range("A2").Value2 = WorksheetFunction.HypGeomDist(1, 2, 3, 5)
}
```

```javascript
/*本示例为B1等单元格分别赋值，使用 HypGeomDist 方法计算超几何分布。*/
function test() {
    Range("B1").Value2 = 2
    Range("B2").Value2 = 12
    Range("B3").Value2 = 6
    Range("B4").Value2 = 24
    console.log(WorksheetFunction.HypGeomDist(Range("B1").Value2, Range("B2").Value2, Range("B3").Value2, Range("B4").Value2))
}
```
