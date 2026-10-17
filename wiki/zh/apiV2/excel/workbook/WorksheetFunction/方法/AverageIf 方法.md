---
title: "AverageIf (方法)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/WorksheetFunction/%E6%96%B9%E6%B3%95/AverageIf%20%E6%96%B9%E6%B3%95.html
updated: 1789700916000
hash: BoFv7Hmf
---

# AverageIf (方法)

返回区域内满足给定条件的所有单元格的平均值（算术平均值）。

## 说明

- range 内包含 TRUE 或 FALSE 的单元格将被忽略。
- 如果 range 或 average_range 中的某个单元格为空单元格，则 AverageIf 会将其忽略。
- 如果条件中的单元格为空，则 AverageIf 会将其作为 0 值处理。
- 如果区域中没有满足条件的单元格，则 AverageIf 将生成一个错误值。
- 可以在条件中使用通配符，包括问号 (?) 和星号 (*)。问号可匹配任意的单个字符；星号可匹配任意一串字符。如果要查找实际的问号或星号，则请在该字符前键入一个波形符 (~)。
- Average_range 的大小和形状不必与 range 相同。求其平均值的实际单元格的确定方法如下：使用 average_range 中左上角的单元格作为起始单元格，然后将与 range 的大小和形状对应的所有单元格包含到其中。例如： | 如果 range 为 | 并且 average_range 为 | 则实际纳入计算的单元格为 | | ---------- | ------------------- | ------------ | | A1:A5 | B1:B5 | B1:B5 | | A1:A5 | B1:B3 | B1:B5 | | A1:B4 | C1:D4 | C1:D4 | | A1:B4 | C1:C2 | C1:D4 |

AverageIf 方法衡量趋中性，趋中性是统计分布中一组数字的中心位置。三种最常见的趋中性衡量方式为：

- **平均值**即算术平均值，是通过将一组数字相加，然后除以数字个数得到的。例如，2、3、3、5、7 和 10 的平均值为 30 除以 6，即为 5。
- **中值**是一组数字中居于中间的数，即在这组数字中，有一半的数字的值比它大，有一半的数字的值比它小。例如，2、3、3、5、7 和 10 的中值为 4。
- **众值**即在一组数字中出现次数最多的数字。例如，2、3、3、5、7 和 10 的众值为 3。

对于对称分布的一组数字，这三种趋中性衡量方式完全相同。对于偏态分布的一组数字，这些衡量方式可能会不同。

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| Arg1 | Range | 必填 | 要求其平均值的一个或多个单元格。 |
| Arg2 | any | 必填 | 定义将对哪些单元格求平均值的条件，其形式可以为数字、表达式、单元格引用或文本。例如，条件可以表示为 32、"32"、">32"、"apples" 或 B4。 |
| Arg3 | any | 可选 | 要求其平均值的实际单元格集合。如果省略，则使用 range。 |

## 返回值

Double

## 示例

```javascript
/*本示例为B3等单元格分别赋值，并使用AverageIf方法计算区域内满足给定条件的所有单元格的平均值。*/
function test() {
    Range("B3").Value2 = "apple"
    Range("B4").Value2 = "apple"
    Range("B5").Value2 = "pear"
    Range("B6").Value2 = "apple"
    Range("C3").Value2 = 20
    Range("C4").Value2 = 40
    Range("C5").Value2 = 60
    Range("C6").Value2 = 50.55
    let averageif1 = Application.WorksheetFunction.AverageIf(Range("B3:B6"), "apple", Range("C3:C6"))
    console.log(averageif1)
}
```

```javascript
/*本示例为A2等单元格分别赋值，使用AverageIf方法计算区域内满足给定条件的所有单元格的平均值，并将结果分别赋值于F2等单元格。*/
function test() {
    Range("A2").Value2 = 12
    Range("A3").Value2 = 12
    Range("A4").Value2 = 12.5
    Range("A5").Value2 = 48
    Range("A6").Value2 = 79
    Range("F2").Value2 = Application.WorksheetFunction.AverageIf(Range("A2:A6"), 12)
    Range("F3").Value2 = Application.WorksheetFunction.AverageIf(Range("A2:A6"), ">12")
}
```
