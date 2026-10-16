---
title: "Erf 方法"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/WorksheetFunction/%E6%96%B9%E6%B3%95/Erf%20%E6%96%B9%E6%B3%95.html
---
# Erf (方法)

返回误差函数在 lower_limit 和 upper_limit 之间的积分。

## 说明

- 如果 lower_limit 是非数值型，则 ERF 将返回错误值 #VALUE!。
- 如果 lower_limit 是负数，则 ERF 将返回错误值 #NUM!。
- 如果 upper_limit 是非数值型，则 ERF 将返回错误值 #VALUE!。
- 如果 upper_limit 是负数，则 ERF 将返回错误值 #NUM!。

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| Arg1 | any | 必填 | Lower_limit - ERF 的积分下限。 |
| Arg2 | any | 可选 | Upper_limit - ERF 的积分上限。如果省略，ERF 将在零到下限之间进行积分。 |

## 返回值

Double

## 示例

```javascript
/*本示例分别赋值A1等单元格，使用 Erf 方法计算误差函数在A1单元格数值和A2单元格数值之间的积分。*/
function test() {
    Range("A1").Value2 = 2
    Range("A2").Value2 = 0.5
    console.log(WorksheetFunction.Erf(Range("A1").Value2, Range("A2").Value2))
}
```

```javascript
/*本示例使用 Erf 方法计算误差函数在参数1和参数2之间的积分，并分别赋值给B1等单元格。*/
function test() {
    Range("B1").Value2 = WorksheetFunction.Erf(0, 0)
    Range("B2").Value2 = WorksheetFunction.Erf(-1, 0.8)
}
```
