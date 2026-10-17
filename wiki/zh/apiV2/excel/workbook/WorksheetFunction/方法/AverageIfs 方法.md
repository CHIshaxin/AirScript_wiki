---
title: "AverageIfs (方法)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/WorksheetFunction/%E6%96%B9%E6%B3%95/AverageIfs%20%E6%96%B9%E6%B3%95.html
updated: 1789700916000
hash: B-Rx3k5d
---

# AverageIfs (方法)

返回满足多个条件的所有单元格的平均值（算术平均值）。

## 说明

- 如果 average_range 中的某个单元格为空单元格，则 AverageIfs 会将其忽略。
- 如果条件区域中的某个单元格为空，则 AverageIfs 会将其作为 0 值处理。
- range 中包含 TRUE 的单元格作为 1 计算；range 中包含 FALSE 的单元格作为 0（零）计算。
- 仅在每个单元格中指定的对应条件都为 True 时，才会在平均值计算过程中使用 average_range 中的该单元格。
- 如果 average_range 中的单元格均为空，或包含无法转换为数字的文本值，则 AverageIfs 将生成一个错误。
- 如果没有满足所有条件的单元格，则 AverageIfs 将生成一个错误值。
- 可以在条件中使用通配符，包括问号 (?) 和星号 (*)。问号可匹配任意的单个字符；星号可匹配任意一串字符。如果要查找实际的问号或星号，则请在该字符前键入一个波形符 (~)。
- 每个 criteria_range 的大小和形状不必与 average_range 相同。计算其平均值的实际单元格的确定方法如下：使用 criteria_range 中左上角的单元格作为起始单元格，然后将与 range 的大小和形状对应的所有单元格包含到其中。例如： | 如果 average_range 为 | 并且 criteria_range 为 | 则实际纳入计算的单元格为 | | ------------------- | -------------------- | ------------ | | A1:A5 | B1:B5 | B1:B5 | | A1:A5 | B1:B3 | B1:B5 | | A1:B4 | C1:D4 | C1:D4 | | A1:B4 | C1:C2 | C1:D4 |

AverageIfs 函数衡量趋中性，趋中性是统计分布中一组数字的中心位置。三种最常见的趋中性衡量方式为：

- **平均值**即算术平均值，是通过将一组数字相加，然后除以数字个数得到的。例如，2、3、3、5、7 和 10 的平均值为 30 除以 6，即为 5。
- **中值**是一组数字中居于中间的数，即在这组数字中，有一半的数字的值比它大，有一半的数字的值比它小。例如，2、3、3、5、7 和 10 的中值为 4。
- **众值**即在一组数字中出现次数最多的数字。例如，2、3、3、5、7 和 10 的众值为 3。

对于对称分布的一组数字，这三种趋中性衡量方式完全相同。对于偏态分布的一组数字，这些衡量方式可能会不同。

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| Arg1 | Range | 必填 | Arg30 - 在其中计算相关条件的一个或多个区域。 |

## 返回值

Double

## 示例

```javascript
/*本示例为D2等单元格分别赋值，使用AverageIfs方法计算满足多个条件的所有单元格的平均值，并将结果分别赋值于L2等单元格。*/
function test() {
    Range("D2").Value2 = "足球"
    Range("D3").Value2 = "篮球"
    Range("D4").Value2 = "足球"
    Range("D5").Value2 = "足球"
    Range("D6").Value2 = "羽毛球"
    Range("E2").Value2 = 70
    Range("E3").Value2 = 80
    Range("E4").Value2 = 100
    Range("E5").Value2 = 130
    Range("E6").Value2 = 5
    Range("F2").Value2 = 1200
    Range("F3").Value2 = 2100
    Range("F4").Value2 = 3300
    Range("F5").Value2 = 4444
    Range("F6").Value2 = 3100
    Range("L2").Value2 = Application.WorksheetFunction.AverageIfs(Range("F2:F6"), Range("D2:D6"), "足球", Range("E2:E6"), ">80")
    Range("L3").Value2 = Application.WorksheetFunction.AverageIfs(Range("E2:E6"), Range("D2:D6"), "足球", Range("F2:F6"), ">1000")
}
```

```javascript
/*本示例为A2等单元格分别赋值，并使用AverageIfs方法计算满足多个条件的所有单元格的平均值。*/
function test() {
    Range("A2").Value2 = "高三"
    Range("A3").Value2 = "高二"
    Range("A4").Value2 = "高三"
    Range("A5").Value2 = "高二"
    Range("A6").Value2 = "高三"
    Range("B2").Value2 = "一班"
    Range("B3").Value2 = "一班"
    Range("B4").Value2 = "二班"
    Range("B5").Value2 = "二班"
    Range("B6").Value2 = "一班"
    Range("C2").Value2 = 98
    Range("C3").Value2 = 97
    Range("C4").Value2 = 94
    Range("C5").Value2 = 95
    Range("C6").Value2 = 99
    let averageifs1 = Application.WorksheetFunction.AverageIfs(Range("C2:C6"), Range("A2:A6"), "高三", Range("B2:B6"), "一班")
    console.log(averageifs1)
}
```
