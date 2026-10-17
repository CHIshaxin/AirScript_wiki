---
title: "Average (方法)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/WorksheetFunction/%E6%96%B9%E6%B3%95/Average%20%E6%96%B9%E6%B3%95.html
updated: 1789700916000
hash: DG-co5OC
---

# Average (方法)

返回参数的平均值（算术平均值）。

## 说明

- 参数可以是数字，也可以是包含数字的名称、数组或引用。
- 直接键入参数列表的逻辑值和数字的文本表示也包括在内。
- 如果数组或引用参数包含文本、逻辑值或空单元格，则这些值将被忽略；但含有零值的单元格包括在内。
- 如果参数为错误值或不能转换为数字的文本，则将导致错误。
- 如果要对引用中的逻辑值和数字的文本表示进行计算，请使用 AVERAGEA 函数。

Average 方法衡量趋中性，趋中性是统计分布中一组数字的中心位置。三种最常见的趋中性衡量方式为：

- **平均值**即算术平均值，是通过将一组数字相加，然后除以数字个数得到的。例如，2、3、3、5、7 和 10 的平均值为 30 除以 6，即为 5。
- **中值**是一组数字中居于中间的数，即在这组数字中，有一半的数字的值比它大，有一半的数字的值比它小。例如，2、3、3、5、7 和 10 的中值为 4。
- **众值**即在一组数字中出现次数最多的数字。例如，2、3、3、5、7 和 10 的众值为 3。

对于对称分布的一组数字，这三种趋中性衡量方式完全相同。对于偏态分布的一组数字，这些衡量方式可能会不同。

**提示** 对单元格求平均值时，应牢记空单元格与含零值单元格之间的区别，尤其在已经清除了**“视图”**选项卡上的**“零值”**复选框（**“工具”**菜单上的**“选项”**命令）的情况下更应如此。空单元格不计算在内，但会计算含零值的单元格。

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| Arg1 | any | 必填 | Arg30 - 要求其平均值的 1 到 30 个数值参数。 |

## 返回值

Double

## 示例

```javascript
/*本示例为B3等单元格分别赋值，并使用Average方法计算参数的平均值。*/
function test() {
    Range("B3").Value2 = 98
    Range("B4").Value2 = 65.6
    Range("B5").Value2 = -32
    Range("B6").Value2 = 789
    Range("B7").Value2 = -12.3
    Range("B8").Value2 = 0
    let average1 = Application.WorksheetFunction.Average(Range("B3:B8").Value2)
    console.log(average1)
}
```

```javascript
/*本示例使用Average方法计算参数的平均值，并将结果赋值A4等单元格。*/
function test() {
    Range("A4").Value2 = Application.WorksheetFunction.Average(17, -51, 94, 5, 84)
    Range("A5").Value2 = Application.WorksheetFunction.Average(1555, 3444, 7777, 9999)
    Range("A6").Value2 = Application.WorksheetFunction.Average(2, 4, -6.6, 8, 10, 12, 14, 16)
    Range("A7").Value2 = Application.WorksheetFunction.Average(1, 3, 5, 7.77, -9.9, 11)
}
```
