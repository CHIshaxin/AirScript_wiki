---
title: "ChiDist (方法)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/WorksheetFunction/%E6%96%B9%E6%B3%95/ChiDist%20%E6%96%B9%E6%B3%95.html
updated: 1789700916000
hash: CU4S9qlo
---

# ChiDist (方法)

返回 χ2 分布的单尾概率。

## 说明

χ2 分布与 χ2 检验相关。使用 χ2检验可以比较观察值和期望值。

例如，某项遗传学实验可能假设下一代植物将呈现出某一组颜色。通过对观察到的结果和所期望的结果进行比较，可以判定初始假设是否有效。

- 如果两个参数中的任意一个为非数值型，则 ChiDist 将生成一个错误。
- 如果 x 为负数，则 ChiDist 将生成一个错误。
- 如果 degrees_freedom 不是整数，则将被截尾取整。
- 如果 degrees_freedom < 1 或 degrees_freedom > 10^10，则 ChiDist 将生成一个错误。
- ChiDist 按 ChiDist =P(X>x) 计算，其中 X 为 χ2 随机变量。

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| Arg1 | double | 必填 | 用来计算分布的值。 |
| Arg2 | double | 必填 | 自由度数。 |

## 返回值

Double

## 示例

```javascript
/*本示例为A2等单元格分别赋值，并使用ChiDist方法返回χ2分布的单尾概率。*/
function test() {
    Range("A2").Value2 = 10
    Range("A3").Value2 = 20
    Range("A4").Value2 = 18.307
    Range("A5").Value2 = 11.4
    let chidist1 = Application.WorksheetFunction.ChiDist(Range("A4").Value2, Range("A2").Value2)
    let chidist2 = Application.WorksheetFunction.ChiDist(Range("A5").Value2, Range("A3").Value2)
    console.log(chidist1)
    console.log(chidist2)
}
```

```javascript
/*本示例使用ChiDist方法返回χ2分布的单尾概率，并将结果赋值于C2等单元格。*/
function test() {
    Range("C2").Value2 = Application.WorksheetFunction.ChiDist(10, 5)
    Range("C3").Value2 = Application.WorksheetFunction.ChiDist(13.5, 30)
    Range("C4").Value2 = Application.WorksheetFunction.ChiDist(8, 15)
}
```
