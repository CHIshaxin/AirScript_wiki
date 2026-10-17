---
title: "Power (方法)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/WorksheetFunction/%E6%96%B9%E6%B3%95/Power%20%E6%96%B9%E6%B3%95.html
updated: 1789700916000
hash: CU8W7QO7
---

# Power (方法)

返回某数的乘幂结果。

## 说明

可以用“^”运算符代替函数 POWER 来表示对底数乘方的幂次，例如 5^2。

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| Arg1 | double | 必填 | Number - 底数，可以为任意实数。 |
| Arg2 | double | 必填 | Power - 指数，底数按该指数次幂乘方。 |

## 返回值

Double

## 示例

```javascript
/*本示例使用 Power 方法分别计算“2的3次方”和“3的4次方”的乘幂结果，并分别赋值给C1和C2单元格。*/
function test() {
    Range("C1").Value2 = WorksheetFunction.Power(2, 3)
    Range("C2").Value2 = WorksheetFunction.Power(3, 4)
}
```

```javascript
/*本示例为A1等单元格分别赋值，使用 Power 方法计算A1单元格和B1单元格的乘幂结果。*/
function test() {
    Range("A1").Value2 = 5
    Range("B1").Value2 = 3
    console.log(WorksheetFunction.Power(Range("A1").Value2, Range("B1").Value2))
}
```
