---
title: "Intercept (方法)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/WorksheetFunction/%E6%96%B9%E6%B3%95/Intercept%20%E6%96%B9%E6%B3%95.html
updated: 1789700916000
hash: CZ6oGtVg
---

# Intercept (方法)

利用现有的 x 值与 y 值计算直线与 y 轴的截距。截距为穿过已知的 known x 和 known y 数据点的线性回归线与 y 轴的交点。当自变量为 0（零）时，使用 INTERCEPT 函数可以决定因变量的值。例如，当所有的数据点都是在室温或更高的温度下取得的，可以用 INTERCEPT 函数预测在 0°C 时金属的电阻。

## 说明

- 参数应为数字，或者是包含数字的名称、数组或引用。
- 如果数组或引用参数包含文本、逻辑值或空单元格，则这些值将被忽略；但含有零值的单元格包括在内。
- 如果 known_y's 和 known_x's 包含的数据点个数不相等或不包含数据点，INTERCEPT 将返回错误值 #N/A。
- INTERCEPT 和 SLOPE 函数中使用的基础算法与 LINEST 函数中使用的基础算法不同。这些算法间的差异可能在共线数据不确定时导致不同的结果。例如，如果 known_y's 参数的数据点个数为 0，而 known_x's 参数的数据点个数为 1：
- INTERCEPT 和 SLOPE 将返回 #DIV/0! 错误。INTERCEPT 和 SLOPE 算法用于寻求一个且只寻求一个答案，但此时可能会有多个答案。
- LINEST 将返回值 0。LINEST 算法用于返回共线数据的合理结构，此时至少可以找到一个答案。

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| Arg1 | any | 必填 | Known_y's - 相关观察值或数据的集合。 |
| Arg2 | any | 必填 | Known_x's - 独立观察值或数据的集合。 |

## 返回值

Double

## 示例

```javascript
/*本示例使用 Intercept 方法利用现有的 x 值与 y 值计算直线与 y 轴的截距，并分别赋值给A2和A3单元格。*/
function test() {
    let arr1 = [1.3, 2.6, 3.56]
    let arr2 = [50, 42, 36]
    Range("A2").Value2 = WorksheetFunction.Intercept(arr2, arr1)
    Range("A3").Value2 = WorksheetFunction.Intercept(arr1, arr2)
}
```

```javascript
/*本示例为D1等单元格分别赋值，使用 Intercept 方法利用现有的 x 值与 y 值计算直线与 y 轴的截距。*/
function test() {
    Range("D1").Value2 = 11
    Range("D2").Value2 = 21
    Range("D3").Value2 = 31
    Range("E1").Value2 = 2
    Range("E2").Value2 = 3
    Range("E3").Value2 = 5
    console.log(WorksheetFunction.Intercept(Range("D1:D3"), Range("E1:E3")))
}
```
