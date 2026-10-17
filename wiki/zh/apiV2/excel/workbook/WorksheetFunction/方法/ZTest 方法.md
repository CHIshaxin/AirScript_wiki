---
title: "ZTest (方法)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/WorksheetFunction/%E6%96%B9%E6%B3%95/ZTest%20%E6%96%B9%E6%B3%95.html
updated: 1789700916000
hash: DSvUdUd1
---

# ZTest (方法)

返回 z 检验的单尾概率值。对于给定的假设总体平均值，ZTEST 返回样本平均值大于数据集（数组）中观察平均值（即观察样本平均值）的概率。

## 说明

- 如果 array 为空，则 ZTEST 将返回错误值 #N/A。
- ZTEST 表示当基础总体平均值为 μ0 时，样本平均值大于观察值 AVERAGE(array) 的概率。由于正态分布是对称的，如果 AVERAGE(array) < μ0，则 ZTEST 的返回值将大于 0.5。
- 当基础总体平均值为 μ0，而样本平均值从 μ0（沿任一方向）变化到 AVERAGE(array) 时，下面的 ET 公式可用于计算双尾概率： =2 * MIN(ZTEST(array,μ0,sigma), 1 - ZTEST(array,μ0,sigma))。

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| Arg1 | any | 必填 | array 为要针对其检验假设总体平均值的数组和数据区域。 |
| Arg2 | double | 必填 | 要检验的值。 |
| Arg3 | any | 可选 | sigma - 总体（已知）的标准偏差。如果省略，则使用样本标准偏差。 |

## 返回值

Double

## 示例

```javascript
/*本示例演示为一组数据计算 z 检验的单尾概率值分别赋值到 A4 单元格和 A5 单元格。*/
function test() {
    let arg1 = [2, 4, 6, 10]
    let arg2 = [20.3, 10.4, 2.6, 10]
    Range("A4").Value2 = WorksheetFunction.ZTest(arg1, 10)
    Range("A5").Value2 = WorksheetFunction.ZTest(arg2, 20, 6.290618014154094)
}
```

```javascript
/*本示例演示为C2等单元格赋值，计算 z 检验的单尾概率值。*/
function test() {
    Range("C2").Value2 = 100
    Range("C3").Value2 = 200
    Range("C4").Value2 = 400
    Range("C5").Value2 = 700
    console.log(WorksheetFunction.ZTest(Range("C2:C5"), 10))
}
```
