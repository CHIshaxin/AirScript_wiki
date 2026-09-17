---
title: "ImCos 方法"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/WorksheetFunction/%E6%96%B9%E6%B3%95/ImCos%20%E6%96%B9%E6%B3%95.html
---
# ImCos (方法)

返回以 x + yi 或 x + yj 文本格式表示的复数的余弦值。

## 说明

- 使用 COMPLEX 可以将实系数和虚系数转换为复数。
- 如果 inumber 为逻辑值，则 IMCOS 将返回错误值 #VALUE!。

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| Arg1 | any | 必填 | Inumber - 要计算其余弦值的复数。 |

## 返回值

String

## 示例

```javascript
/*本示例使用 ImCos 方法计算以 x + yi 或 x + yj 文本格式表示的复数的余弦值，并分别赋值给D1和D2单元格。*/
function test() {
    Range("D1").Value2 = WorksheetFunction.ImCos(WorksheetFunction.Complex(20, 6))
    Range("D2").Value2 = WorksheetFunction.ImCos("1+3j")
}
```

```javascript
/*本示例为D3等单元格分别赋值，使用 ImCos 方法计算以 x + yi 文本格式表示的D3单元格中复数的余弦值。*/
function test() {
    Range("D3").Formula2 = "=COMPLEX(1,4)"
    console.log(WorksheetFunction.ImCos(Range("D3").Value2))
}
```
