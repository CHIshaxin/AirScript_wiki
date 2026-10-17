---
title: "Irr (方法)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/WorksheetFunction/%E6%96%B9%E6%B3%95/Irr%20%E6%96%B9%E6%B3%95.html
updated: 1789700916000
hash: dOSWSekE
---

# Irr (方法)

返回由数值代表的一组现金流的内部收益率。这些现金流不必为均衡的，但作为年金，它们必须按固定的间隔产生，如按月或按年。内部收益率为投资的回收利率，其中包含定期支付（负值）和定期收入（正值）。

## 说明

- Values 必须至少包含一个正值和一个负值，以计算内部报酬率。
- IRR 根据值的顺序来解释现金流的顺序。请确保按照所需的顺序输入了支出和收入值。
- 如果数组或引用参数包含文本、逻辑值或空单元格，则这些值将被忽略。
- ET 使用迭代技术计算 IRR。从 guess 开始，IRR 进行循环计算，直至结果的精确度达到 0.00001%。如果 IRR 经过 20 次尝试后仍未找到结果，则返回错误值 #NUM!。
- 在大多数情况下，并不需要为 IRR 的计算提供 guess 值。如果省略 guess，则假设它为 0.1 (10%)。
- 如果 IRR 给出错误值 #NUM!，或结果并不接近期望值，则请用不同的 guess 值再试一次。

函数 IRR 与函数 NPV（净现值函数）的关系十分密切。函数 IRR 计算出的收益率即是净现值为 0（零）时的利率。下面的公式显示了函数 NPV 和函数 IRR 的相互关系：

`NPV(IRR(B1:B6),B1:B6)` 等于 3.60E-08 [在 IRR 计算的精确度要求之内，数值 3.60E-08 可以当作 0（零）的有效值]。

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| Arg1 | any | 必填 | Values - 一个数组，或者是对包含用来计算内部报酬率的数字的单元格的引用。 |
| Arg2 | any | 可选 | Guess - 一个认为接近 IRR 结果的数字。 |

## 返回值

Double

## 示例

```javascript
/*本示例使用 Irr 方法计算由数组的数字表示的一系列现金流的内部收益率，并分别赋值给A1和A2单元格。*/
function test() {
    Range("A1").Value2 = WorksheetFunction.Irr([1.1, 2.2, 3.3, 5.6, -2.6])
    Range("A2").Value2 = WorksheetFunction.Irr([-4, -5, -6, -8, 11], 10000)
}
```

```javascript
/*本示例为B1等单元格分别赋值，使用 Irr 方法计算由区域B1:B4的数字表示的一系列现金流的内部收益率。*/
function test() {
    Range("B1").Value2 = 123
    Range("B2").Value2 = 234
    Range("B3").Value2 = 345
    Range("B4").Value2 = -567
    console.log(WorksheetFunction.Irr(Range("B1:B4")))
}
```
