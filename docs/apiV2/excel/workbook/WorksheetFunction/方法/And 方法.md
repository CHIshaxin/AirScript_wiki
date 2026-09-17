---
title: "And 方法"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/WorksheetFunction/%E6%96%B9%E6%B3%95/And%20%E6%96%B9%E6%B3%95.html
---
# And (方法)

如果其所有参数都为 TRUE，则返回 TRUE；如果一个或多个参数为 FALSE，则返回 FALSE。

## 说明

- 参数的计算结果必须为逻辑值（如 TRUE 或 FALSE），或者参数必须是包含逻辑值的数组或引用。
- 如果数组或引用参数中包含文本或空单元格，则这些值将被忽略。
- 如果指定区域内不包含逻辑值，则此方法将生成一个错误值。

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| Arg1 | any | 必填 | Arg30 - 1 到 30 个需要进行检验的条件，可为 TRUE 或 FALSE。 |

## 返回值

Boolean

## 示例

```javascript
/*本示例使用And计算所有参数结果，并将结果赋值于E5等单元格。*/
function test() {
    Range("E5").Value2 = Application.WorksheetFunction.And(true, true)
    Range("E6").Value2 = Application.WorksheetFunction.And(false, false)
    Range("E7").Value2 = Application.WorksheetFunction.And(false, true)
}
```

```javascript
/*本示例为H5等单元格分别赋值，并使用And计算所有参数结果。*/
function test() {
    Range("H5").Value2 = 78
    Range("H6").Value2 = 56
    let and1 = Application.WorksheetFunction.And(Range("H5").Value2 > Range("H6").Value2)
    let and2 = Application.WorksheetFunction.And(true, Range("H5").Value2 < Range("H6").Value2, true)
    let and3 = Application.WorksheetFunction.And(Range("H5").Value2 > Range("H6").Value2, false, true)
    console.log(and1)
    console.log(and2)
    console.log(and3)
}
```
