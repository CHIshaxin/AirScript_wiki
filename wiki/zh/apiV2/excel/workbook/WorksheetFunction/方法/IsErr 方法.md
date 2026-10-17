---
title: "IsErr (方法)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/WorksheetFunction/%E6%96%B9%E6%B3%95/IsErr%20%E6%96%B9%E6%B3%95.html
updated: 1789700916000
hash: D5w5nPgM
---

# IsErr (方法)

检查值类型，并根据值是否引用 #N/A 之外的任何错误值而返回 TRUE 或 FALSE。

## 说明

- IS 函数的值参数不能进行转换。例如，在其他需要数字的大多数函数中，文本值 "19" 会被转换成数字 19。然而在公式 ISNUMBER("19") 中，"19" 并未从文本值转换成数字，因此函数 ISNUMBER 返回 FALSE。
- IS 函数在用公式检验计算结果时十分有用。 与 IF 函数结合使用时，它们提供用于在公式中查找错误的方法。

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| Arg1 | any | 必填 | Value - 要检验的值。可以是空白（空单元格）、错误值、逻辑值、文本值、数字、引用值或对以上任意要检验的值进行引用的名称。 |

## 返回值

Boolean

## 示例

```javascript
/*本示例为A2单元格赋值，使用 IsErr 方法判断值是否引用除 #N/A 以外的任何错误值，并分别赋值给D7和D8单元格。*/
function test() {
    Range("A2").Formula2 = "=NA()"
    Range("D7").Value2 = WorksheetFunction.IsErr(Range("A3").Value2)
    Range("D8").Value2 = WorksheetFunction.IsErr(Range("A2"))
}
```

```javascript
/*本示例使用 IsErr 方法判断B3单元格是否引用除 #N/A 以外的任何错误值。*/
function test() {
    console.log(WorksheetFunction.IsErr(Range("B3").Value2))
}
```
