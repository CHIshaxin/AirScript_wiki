---
title: "PercentRank 方法"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/WorksheetFunction/%E6%96%B9%E6%B3%95/PercentRank%20%E6%96%B9%E6%B3%95.html
---
# PercentRank (方法)

返回特定数值在一个数据集中的百分比排位。此函数可用于计算某个值在数据集中所处的位置。例如，可以使用函数 PERCENTRANK 计算某个特定的能力测试得分在所有的能力测试得分中的位置。

## 说明

- 如果数组为空，函数 PERCENTRANK 返回错误值 #NUM!。
- 如果 significance < 1，函数 PERCENTRANK 返回错误值 #NUM!。
- 如果数组里没有与 x 相匹配的值，函数 PERCENTRANK 将进行插值以返回正确的百分比排位。

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| Arg1 | any | 必填 | array - 所含数值定义相对位置的数组或数据区域。 |
| Arg2 | double | 必填 | x - 想知道其排位的值。 |
| Arg3 | any | 可选 | significance - 一个可选值，表示返回的百分数值的有效位数。如果省略，函数 PERCENTRANK 保留三位小数 (0.xxx)。 |

## 返回值

Double

## 示例

```javascript
/*本示例使用 PercentRank 方法分别计算数值4在数组arr1的百分比排位、数值10在数组arr1的百分比排位（保留三位），并将结果分别赋值给D1和D2单元格。*/
function test() {
    let arr1 = [1, 2, 3.2, 5, 6, 10.6]
    Range("D1").Value2 = WorksheetFunction.PercentRank(arr1, 4)
    Range("D2").Value2 = WorksheetFunction.PercentRank(arr1, 10, 3)
}
```

```javascript
/*本示例为A1等单元格分别赋值，使用 PercentRank 方法计算单元格B1中的数值在区域A1:A5的百分比排位。*/
function test() {
    Range("A1").Value2 = 12
    Range("A2").Value2 = 14
    Range("A3").Value2 = 23
    Range("A4").Value2 = 32
    Range("A5").Value2 = 42
    Range("B1").Value2 = 13
    console.log(WorksheetFunction.PercentRank(Range("A1:A5"), Range("B1").Value2))
}
```
