---
title: "ImSum (方法)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/WorksheetFunction/%E6%96%B9%E6%B3%95/ImSum%20%E6%96%B9%E6%B3%95.html
updated: 1789700916000
hash: CMf1X43_
---

# ImSum (方法)

返回以 x + yi 或 x + yj 文本格式表示的两个或更多个复数的和。

## 说明

- 使用 COMPLEX 可以将实系数和虚系数转换为复数。

## 返回值

String

## 示例

```javascript
/*本示例使用 ImSum 方法计算 x + yi 或 x + yj 文本格式表示的复数和，并分别赋值给D3和D4单元格。*/
function test() {
    Range("D3").Value2 = WorksheetFunction.ImSum(WorksheetFunction.Complex(1, 1), WorksheetFunction.Complex(1, 2))
    Range("D4").Value2 = WorksheetFunction.ImSum("2-2j", "1+3j", "3+5j")
}
```

```javascript
/*本示例为B1等单元格分别赋值，使用 ImSum 方法以 x + yi 文本格式表示两个复数的和。*/
function test() {
    Range("B1").Formula = "=COMPLEX(4,5)"
    Range("B2").Value2 = "4+7i"
    console.log(WorksheetFunction.ImSum(Range("B1").Value2, Range("B2").Value2))
}
```

param Arg1 - Arg30 - Inumber1,inumber2,... - 1 到 29 个需要相加的复数。
