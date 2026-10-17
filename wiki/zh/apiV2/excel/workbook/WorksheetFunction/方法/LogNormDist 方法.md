---
title: "LogNormDist (方法)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/WorksheetFunction/%E6%96%B9%E6%B3%95/LogNormDist%20%E6%96%B9%E6%B3%95.html
updated: 1789700916000
hash: Dj_xKSRN
---

# LogNormDist (方法)

返回 x 的对数累积分布函数，其中 ln(x) 是服从参数 mean 和 standard_dev 的正态分布。使用此函数可以分析经过对数变换的数据。

## 说明

- 如果任一参数是非数字型，则 LOGNORMDIST 返回错误值 #VALUE!。
- 如果 x ≤ 0 或 standard_dev ≤ 0，则 LOGNORMDIST 返回错误值 #NUM!。

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| Arg1 | double | 必填 | x - 用于计算函数的值。 |
| Arg2 | double | 必填 | mean - ln(x) 的平均值。 |
| Arg3 | double | 必填 | standard_dev - ln(x) 的标准偏差。 |

## 返回值

Double

## 示例

```javascript
/*本示例使用 LogNormDist 方法分别计算数值10和100的对数累积分布函数，并分别赋值给C2和C3单元格。*/
function test() {
    Range("C2").Value2 = WorksheetFunction.LogNormDist(10, 21.3, 0.8)
    Range("C3").Value2 = WorksheetFunction.LogNormDist(100, 111, 5)
}
```

```javascript
/*本示例为C1等单元格分别赋值，使用 LogNormDist 方法计算单元格C1的对数累积分布函数。*/
function test() {
    Range("C1").Value2 = 12
    Range("C2").Value2 = 11
    Range("C3").Value2 = 10
    console.log(WorksheetFunction.LogNormDist(Range("C1").Value2, Range("C2").Value2, Range("C3").Value2))
}
```
