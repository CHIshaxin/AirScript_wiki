---
title: "RandBetween (方法)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/WorksheetFunction/%E6%96%B9%E6%B3%95/RandBetween%20%E6%96%B9%E6%B3%95.html
updated: 1789700916000
hash: BA9SPfHg
---

# RandBetween (方法)

返回介于指定的两个数字之间的一个随机整数。每次计算工作表时都将返回一个新的随机整数。

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| Arg1 | any | 必填 | Bottom - RANDBETWEEN 将返回的最小整数。 |
| Arg2 | any | 必填 | Top - RANDBETWEEN 将返回的最大整数。 |

## 返回值

Double

## 示例

```javascript
/*本示例使用 RandBetween 方法分别获取数字100至150之间和数字10至30之间的随机整数，并赋值给B3和B4单元格。*/
function test() {
    Range("B3").Value2 = WorksheetFunction.RandBetween(100, 150)
    Range("B4").Value2 = WorksheetFunction.RandBetween(10, 30)
}
```

```javascript
/*本示例为A1等单元格分别赋值，使用 RandBetween 方法获取A1单元格中数字和A2单元格中数字之间的随机整数。*/
function test() {
    Range("A1").Value2 = 12
    Range("A2").Value2 = 21
    console.log(WorksheetFunction.RandBetween(Range("A1").Value2, Range("A2").Value2))
}
```
