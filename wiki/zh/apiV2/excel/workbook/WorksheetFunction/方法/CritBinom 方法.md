---
title: "CritBinom (方法)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/WorksheetFunction/%E6%96%B9%E6%B3%95/CritBinom%20%E6%96%B9%E6%B3%95.html
updated: 1789700916000
hash: B8uhZV-7
---

# CritBinom (方法)

返回使累积二项式分布大于或等于临界值的最小值。

## 说明

可以将此函数用于质量保证。例如，使用 CritBinom 来决定最多允许出现多少个有缺陷的部件，才可以保证当整批产品在离开装配线时检验合格。

- 如果任一参数为非数值型，则 CritBinom 将生成一个错误。
- 如果 trials 不是整数，则将被截尾取整。
- 如果 trials < 0，则 CritBinom 将生成一个错误。
- 如果 probability_s < 0 或 probability_s > 1，则 CritBinom 将生成一个错误。
- 如果 alpha < 0 或 alpha > 1，则 CritBinom 将生成一个错误。

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| Arg1 | double | 必填 | 贝努利试验次数。 |
| Arg2 | double | 必填 | 每次试验成功的概率。 |
| Arg3 | double | 必填 | 临界值。 |

## 返回值

Double

## 示例

```javascript
/*本示例为D3等单元格分别赋值，并使用CritBinom方法计算累积二项式分布大于等于临界值的最小值。*/
function test() {
    Range("D3").Value2 = 6
    Range("D4").Value2 = 0.5
    Range("D5").Value2 = 0.75
    let critbinom1 = Application.WorksheetFunction.CritBinom(Range("D3").Value2, Range("D4").Value2, Range("D5").Value2)
    console.log(critbinom1)
}
```

```javascript
/*本示例使用CritBinom方法计算累积二项式分布大于等于临界值的最小值，并将结果赋值于A3单元格。*/
function test() {
    Range("A3").Value2 = Application.WorksheetFunction.CritBinom(10, 0.7, 0.5)
    Range("B3").Value2 = Application.WorksheetFunction.CritBinom(20, 0.4, 0.65)
    Range("C3").Value2 = Application.WorksheetFunction.CritBinom(15, 0.9, 0.33)
}
```
