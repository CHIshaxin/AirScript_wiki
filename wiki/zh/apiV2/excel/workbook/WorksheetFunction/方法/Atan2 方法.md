---
title: "Atan2 (方法)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/WorksheetFunction/%E6%96%B9%E6%B3%95/Atan2%20%E6%96%B9%E6%B3%95.html
updated: 1789700916000
hash: B425RM-8
---

# Atan2 (方法)

返回指定的 x 和 y 坐标值的反正切值。反正切值是角度，是从 x 轴到通过原点 (0, 0) 和坐标点 (x_num, y_num) 的直线之间的夹角。该角度用弧度给出，介于 -pi 和 pi 之间（不包括 -pi）。

## 说明

- 结果为正表示从 x 轴逆时针旋转的角度；结果为负表示从 x 轴顺时针旋转的角度。
- Atan2(a,b) 等于 Atan(b/a)，不同之处在于 Atan2 中的 a 可等于 0。
- 如果 x_num 和 y_num 为 0，则 Atan2 将返回一个错误值。
- 要用度表示反正切值，请将结果乘以 180/PI( ) 或使用 Degrees 方法。

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| Arg1 | double | 必填 | 点的 x 坐标。 |
| Arg2 | double | 必填 | 点的 y 坐标。 |

## 返回值

Double

## 示例

```javascript
/*本示例为B3等单元格分别赋值，并使用Atan2计算指定的x和y坐标值的反正切值。*/
function test() {
    Range("B3").Value2 = 1
    Range("B4").Value2 = -2
    Range("B5").Value2 = 3
    Range("B6").Value2 = -4
    let tan1 = Application.WorksheetFunction.Atan2(Range("B3").Value2, Range("B4").Value2)
    let tan2 = Application.WorksheetFunction.Atan2(Range("B3").Value2, Range("B5").Value2)
    let tan3 = Application.WorksheetFunction.Atan2(Range("B4").Value2, Range("B6").Value2)
    let tan4 = Application.WorksheetFunction.Atan2(Range("B6").Value2, Range("B3").Value2)
    console.log(tan1)
    console.log(tan2)
    console.log(tan3)
    console.log(tan4)
}
```

```javascript
/*本示例使用Atan2计算指定的x和y坐标值的反正切值，并将结果分别赋值于C4等单元格*/
function test() {
    Range("C4").Value2 = Application.WorksheetFunction.Atan2(1, 3)
    Range("D4").Value2 = Application.WorksheetFunction.Atan2(4.5, -7)
    Range("E4").Value2 = Application.WorksheetFunction.Atan2(-5, 5.5)
    Range("F4").Value2 = Application.WorksheetFunction.Atan2(-3, 4)
}
```
