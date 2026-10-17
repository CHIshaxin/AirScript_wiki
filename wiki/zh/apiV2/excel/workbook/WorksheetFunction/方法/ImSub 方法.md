---
title: "ImSub (方法)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/WorksheetFunction/%E6%96%B9%E6%B3%95/ImSub%20%E6%96%B9%E6%B3%95.html
updated: 1789700916000
hash: DjXoBecW
---

# ImSub (方法)

返回以 x + yi 或 x + yj 文本格式表示的两个复数的差。

## 说明

- 使用 COMPLEX 可以将实系数和虚系数转换为复数。

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| Arg1 | any | 必填 | Inumber1 - 作为被减数的复数。 |
| Arg2 | any | 必填 | Inumber2 - 作为减数的复数。 |

## 返回值

String

## 示例

```javascript
/*本示例使用 ImSub 方法计算以 x + yi 或 x + yj 文本格式表示的两个复数差，并分别赋值给A2和A3单元格。*/
function test() {
    Range("A2").Value2 = WorksheetFunction.ImSub(WorksheetFunction.Complex(12, 4), WorksheetFunction.Complex(10, 20))
    Range("A3").Value2 = WorksheetFunction.ImSub("1-4j", "6+2j")
}
```

```javascript
/*本示例为C1等单元格分别赋值，使用 ImSub 方法计算以 x + yi文本格式表示单元格C1和C2中的两个复数的差。*/
function test() {
    Range("C1").Formula = "=COMPLEX(3,11)"
    Range("C2").Value2 = "5+6i"
    console.log(WorksheetFunction.ImSub(Range("C1").Value2, Range("C2").Value2))
}
```
