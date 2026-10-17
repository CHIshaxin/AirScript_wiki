---
title: "ImLn (方法)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/WorksheetFunction/%E6%96%B9%E6%B3%95/ImLn%20%E6%96%B9%E6%B3%95.html
updated: 1789700916000
hash: DleLdeZY
---

# ImLn (方法)

返回以 x + yi 或 x + yj 文本格式表示的复数的自然对数。

## 说明

- 使用 COMPLEX 可以将实系数和虚系数转换为复数。

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| Arg1 | any | 必填 | Inumber - 要计算其自然对数的复数。 |

## 返回值

String

## 示例

```javascript
/*本示例使用 ImLn 方法计算以 x + yi 或 x + yj 文本格式表示的复数的自然对数，并分别赋值给E3和E4单元格。*/
function test() {
    Range("E3").Value2 = WorksheetFunction.ImLn(WorksheetFunction.Complex(12, 3))
    Range("E4").Value2 = WorksheetFunction.ImLn("10+10j")
}
```

```javascript
/*本示例使用 ImLn 方法计算以 x + yi 文本格式表示的单元格B3中复数的自然对数。*/
function test() {
    Range("B3").Formula2 = "=COMPLEX(4,6)"
    console.log(WorksheetFunction.ImLn(Range("B3").Value2))
}
```
