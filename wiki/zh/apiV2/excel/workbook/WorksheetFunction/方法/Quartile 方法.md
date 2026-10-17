---
title: "Quartile (方法)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/WorksheetFunction/%E6%96%B9%E6%B3%95/Quartile%20%E6%96%B9%E6%B3%95.html
updated: 1789700916000
hash: qLQWgw3F
---

# Quartile (方法)

返回数据集的四分位数。四分位数通常用于在销售额和调查数据中对总体进行分组。例如，可以使用函数 QUARTILE 求得总体中前 25% 的收入值。

## 说明

| 如果 quart 等于 | QUARTILE 就返回 |
| --- | --- |
| 0 | 最小值 |
| 1 | 第一个四分位数（第 25 个百分点值） |
| 2 | 中值（第 50 个百分点值） |
| 3 | 第三个四分位数（第 75 个百分点值） |
| 4 | 最大值 |

- 如果 array 为空，则 QUARTILE 将返回错误值 #NUM!。
- 如果 quart 不为整数，则将被截尾取整。
- 如果 quart < 0 或 quart > 4，则 QUARTILE 将返回错误值 #NUM!。
- 当 quart 分别等于 0（零）、2 和 4 时，MIN、MEDIAN 和 MAX 返回的值与 QUARTILE 返回的值相同。

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| Arg1 | any | 必填 | array - 要计算其四分位数值的数值数组或单元格区域。 |
| Arg2 | double | 必填 | Quart - 指示要返回哪一个值。 |

## 返回值

Double

## 示例

```javascript
/*本示例使用 Quartile 方法计算数组的四分位数，并分别赋值给A4和A5单元格。*/
function test() {
    let arr1 = [10, 10.2, 3.2, 5, 3.6]
    Range("A4").Value2 = WorksheetFunction.Quartile(arr1, 0)
    Range("A5").Value2 = WorksheetFunction.Quartile(arr1, 2)
}
```

```javascript
/*本示例为B1等单元格分别赋值，使用 Quartile 方法计算表格区域B1:B4的四分位数。*/
function test() {
    Range("B1").Value2 = 12
    Range("B2").Value2 = 24
    Range("B3").Value2 = 36
    Range("B4").Value2 = 48
    console.log(WorksheetFunction.Quartile(Range("B1:B4"), 3))
}
```
