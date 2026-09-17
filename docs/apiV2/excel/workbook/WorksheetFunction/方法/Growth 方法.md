---
title: "Growth 方法"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/WorksheetFunction/%E6%96%B9%E6%B3%95/Growth%20%E6%96%B9%E6%B3%95.html
---
# Growth (方法)

根据现有的数据预测指数增长值。根据现有的 x 值和 y 值，GROWTH 函数返回一组新的 x 值对应的 y 值。可以使用 GROWTH 工作表函数来拟合满足现有 x 值和 y 值的指数曲线。

## 说明

- 如果数组 known_y's 在单独一列中，则 known_x's 的每一列被解释为一个独立的变量。
- 如果数组 known_y's 在单独一行中，则 known_x's 的每一行被解释为一个独立的变量。
- 如果 known_y's 中的任何数字为零或为负数，则 GROWTH 将返回错误值 #NUM!。
- 数组 known_x's 可以包含一组或多组变量。如果只用到一个变量，只要 known_y's 和 known_x's 维数相同，它们就可以是任何形状的区域。如果用到多个变量，则 known_y's 必须为向量（即必须为一行或一列）。
- 如果省略 known_x's，则假定该数组为 {1,2,3,...}，其大小与 known_y's 相同。
- New_x's 与 known_x's 一样，每个独立变量必须为单独的一列（或一行）。因此，如果 known_y's 是单列的，known_x's 和 new_x's 就必须有同样的列数。如果 known_y's 是单行的，known_x's 和 new_x's 就必须有同样的行数。
- 如果省略 new_x's，则假定它和 known_x's 相同。
- 如果 known_x's 与 new_x's 都被省略，则假定它们为数组 array {1,2,3,...}，其大小与 known_y's 相同。
- 如果 const 为 TRUE 或被省略，b 将按正常方式计算。
- 如果 const 为 FALSE，b 将被设置为等于 1，m 值将被调整以满足 y = m^x。
- 对于返回结果为数组的公式，在选择正确的单元格个数后，必须以数组公式的形式输入。
- 当为参数（如 known_x's）输入数组常量时，应当使用逗号分隔同一行中的各个值，用分号分隔不同行中的值。

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| Arg1 | any | 必填 | Known_y's - 关系表达式 y=b\*m^x 中一组已知的 y 值。 |
| Arg2 | any | 可选 | Known_x's - 关系表达式 y=b\*m^x 中一组已知的 x 值，为可选参数。 |
| Arg3 | any | 可选 | New_x's - 想让 GROWTH 为其返回对应 y 值的新 x 值。 |
| Arg4 | any | 可选 | Const - 一个逻辑值，用于指定是否强制常量 b 等于 1。 |

## 返回值

Variant

## 示例

```javascript
/*本示例使用 Growth 方法根据现有的数组arr1和arr2中的数据预测指数增长值，并分别赋值给D1等单元格。*/
function test() {
    let arr1 = [123, 234, 345, 456]
    let arr2 = [3, 6, 9, 15]
    Range("D1:G1").Value2 = WorksheetFunction.Growth(arr1, arr2, 5, false)
    Range("D2:G2").Value2 = WorksheetFunction.Growth(arr2, arr1, 10)
}
```

```javascript
/*本示例为A1等单元格分别赋值，使用 Growth 方法根据现有的区域A1:A2和B1:B2中的数据预测指数增长值。*/
function test() {
    Range("A1").Value2 = 12
    Range("A2").Value2 = 26
    Range("B1").Value2 = 666
    Range("B2").Value2 = 987
    let arr = WorksheetFunction.Growth(Range("A1:A2"), Range("A1:A2"))
    console.log(arr.toString())
}
```
