---
title: "Trend (方法)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/WorksheetFunction/%E6%96%B9%E6%B3%95/Trend%20%E6%96%B9%E6%B3%95.html
updated: 1789700916000
hash: DOWzEWv3
---

# Trend (方法)

返回一条线性回归拟合线的一组纵坐标值（y 值）。即找到适合给定的数组 known_y's 和 known_x's 的直线（用最小二乘法），并返回指定数组 new_x's 值在直线上对应的 y 值。

## 说明

- 如果数组 known_y's 在单独一列中，则 known_x's 的每一列被解释为一个独立的变量。
- 如果数组 known_y's 在单独一行中，则 known_x's 的每一行被解释为一个独立的变量。
- 数组 known_x's 可以包含一组或多组变量。如果只用到一个变量，只要 known_y's 和 known_x's 维数相同，它们就可以是任何形状的区域。如果用到多个变量，则 known_y's 必须为向量（即必须为一行或一列）。
- 如果省略 known_x's，则假定该数组为 {1,2,3,...}，其大小与 known_y's 相同。
- New_x's 与 known_x's 一样，每个独立变量必须为单独的一列（或一行）。因此，如果 known_y's 是单列的，known_x's 和 new_x's 就必须有同样的列数。如果 known_y's 是单行的，known_x's 和 new_x's 就必须有同样的行数。
- 如果省略 new_x's，则将假定它和 known_x's 一样。
- 如果 known_x's 和 new_x's 都省略，则将假定它们为数组 {1,2,3,...}，大小与 known_y's 相同。
- 如果 const 为 TRUE 或被省略，b 将按正常方式计算。
- 如果 const 为 FALSE，b 将被设置为等于 0（零），并将对 m 值进行调整，以使 y = mx。
- 有关 ET 对数据进行直线拟合的信息，请参阅 LINEST。
- 可以使用 TREND 计算同一变量的不同乘方的回归值来拟合多项式曲线。例如，假设 A 列包含 y 值，B 列包含 x 值。可以在 C 列中输入 x^2，在 D 列中输入 x^3，等等，然后根据 A 列对 B 列到 D 列进行回归计算。
- 对于返回结果为数组的公式，必须以数组公式的形式输入。
- 当为参数（如 known_x's）输入数组常量时，应当使用逗号分隔同一行中的各个值，用分号分隔不同行中的值。

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| Arg1 | any | 必填 | Known_y's - 关系表达式 y = mx + b 中已知的 y 值集合。 |
| Arg2 | any | 可选 | Known_x's - 关系表达式 y = mx + b 中可能已知的可选 x 值集合。 |
| Arg3 | any | 可选 | New_x's - 想让 TREND 为其返回对应 y 值的新 x 值。 |
| Arg4 | any | 可选 | Const - 一个逻辑值，用于指定是否强制常量 b 等于 0。 |

## 返回值

Variant

## 示例

```javascript
/*本示例使用 Trend 方法计算沿线性趋势的值，并分别赋值到A1等单元格。*/
function test() {
    let arr1 = [10, 20, 30, 40, 50]
    let arr2 = [1, 2.3, 3, 4.5, 5]
    let arr3 = [6, -7, 8]
    Range("A1:C1").Value2 = WorksheetFunction.Trend(arr1, arr2, arr3, true)
    Range("A2:C2").Value2 = WorksheetFunction.Trend(arr1, arr2, arr3, false)
}
```

```javascript
/*本示例为C2等单元格分别赋值，使用 Trend 方法计算沿线性趋势的值。*/
function test() {
    Range("C2").Value2 = 10
    Range("C3").Value2 = 2
    Range("C4").Value2 = 3
    Range("C5").Value2 = 400
    Range("C6").Value2 = 5
    Range("D2").Value2 = 1
    Range("D3").Value2 = 2.5
    Range("D4").Value2 = 3.2
    Range("D5").Value2 = 4
    Range("D6").Value2 = 5
    Range("E2").Value2 = 6
    Range("E3").Value2 = -7
    Range("E4").Value2 = 8
    let trend = WorksheetFunction.Trend(Range("C2:C6"), Range("D2:D6"), Range("E2:E4"), false)
    console.log(trend.toString())
}
```
