---
title: "RSq (方法)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/WorksheetFunction/%E6%96%B9%E6%B3%95/RSq%20%E6%96%B9%E6%B3%95.html
updated: 1789700916000
hash: JCEW9xf2
---

# RSq (方法)

返回根据 known_y's 和 known_x's 中数据点计算得出的 Pearson 乘积矩相关系数的平方。有关详细信息，请参阅函数 PEARSON。R 平方值可以解释为 y 方差与 x 方差的比例。

## 说明

- 参数可以是数字，也可以是包含数字的名称、数组或引用。
- 直接键入参数列表的逻辑值和数字的文本表示也包括在内。
- 如果数组或引用参数包含文本、逻辑值或空单元格，则这些值将被忽略；但含有零值的单元格包括在内。
- 如果参数为错误值或不能转换为数字的文本，则将导致错误。
- 如果 known_y's 和 known_x's 为空或其数据点个数不同，则 RSQ 将返回错误值 #N/A。
- 如果 known_y's 和 known_x's 仅包含 1 个数据点，则 RSQ 将返回错误值 #DIV/0!。 RSQ 返回 r2，即此相关系数的平方。

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| Arg1 | any | 必填 | Known_y's - 数据点数组或区域。 |
| Arg2 | any | 必填 | Known_x's - 数据点数组或区域。 |

## 返回值

Double

## 示例

```javascript
/*本示例使用 RSq 方法根据数组arr1、arr2和arr3中数据点计算得出的 Pearson 乘积矩相关系数的平方，并分别赋值给A1和A2单元格。*/
function test() {
    let arr1 = [1, 2, 3.2, 5]
    let arr2 = [-0.5, 4, 6, 8.6]
    let arr3 = [1.2, 3, 2.6, 8]
    Range("A1").Value2 = WorksheetFunction.RSq(arr1, arr2)
    Range("A2").Value2 = WorksheetFunction.RSq(arr3, arr2)
}
```

```javascript
/*本示例为A1等单元格分别赋值，使用 RSq 方法根据区域A1:A3和区域B1:B3数据点计算得出的 Pearson 乘积矩相关系数的平方。*/
function test() {
    Range("A1").Value2 = 1.3
    Range("A2").Value2 = 3.6
    Range("A3").Value2 = 4.5
    Range("B1").Value2 = 4.6
    Range("B2").Value2 = 7
    Range("B3").Value2 = 5
    console.log(WorksheetFunction.RSq(Range("A1:A3"), Range("B1:B3")))
}
```
