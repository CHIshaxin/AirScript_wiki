---
title: "Weibull (方法)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/WorksheetFunction/%E6%96%B9%E6%B3%95/Weibull%20%E6%96%B9%E6%B3%95.html
updated: 1789700916000
hash: D6C_gqWj
---

# Weibull (方法)

返回韦伯分布。使用此分布可以进行可靠性分析，例如计算设备失效的平均时间。

## 说明

- 如果 x、alpha 或 beta 为非数值型，则 WEIBULL 将返回错误值 #VALUE!。
- 如果 x < 0，则 WEIBULL 将返回错误值 #NUM!。
- 如果 alpha ≤ 0 或 beta ≤ 0，则 WEIBULL 将返回错误值 #NUM!。

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| Arg1 | double | 必填 | x - 用于计算函数的值。 |
| Arg2 | double | 必填 | alpha - 分布参数。 |
| Arg3 | double | 必填 | beta - 分布参数。 |
| Arg4 | boolean | 必填 | cumulative - 决定函数的形式。 |

## 返回值

Double

## 示例

```javascript
/*本示例演示用 Weibull 计算韦伯分布。*/
function test() {
    Range("A3").Value2 = WorksheetFunction.Weibull(1, 2, 2, true)
    Range("A4").Value2 = WorksheetFunction.Weibull(2, 4, 1, false)
}
```

```javascript
/*本示例演示为B3等单元格赋值，再用 Weibull 计算韦伯分布。*/
function test() {
    Range("B3").Value2 = 2.3
    Range("B4").Value2 = 3
    Range("B5").Value2 = 5.62
    console.log(WorksheetFunction.Weibull(Range("B3").Value2, Range("B4").Value2, Range("B5").Value2, true))
}
```
