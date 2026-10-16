---
title: "IsLogical 方法"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/WorksheetFunction/%E6%96%B9%E6%B3%95/IsLogical%20%E6%96%B9%E6%B3%95.html
---
# IsLogical (方法)

检查值类型，并根据值是否引用逻辑值而返回 TRUE 或 FALSE。

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
/*本示例使用 IsLogical 方法判断值的类型是否为逻辑值，并分别赋值给E1和E2单元格。*/
function test() {
    Range("E1").Value2 = WorksheetFunction.IsLogical(true)
    Range("E2").Value2 = WorksheetFunction.IsLogical(55)
}
```

```javascript
/*本示例使用 IsLogical 方法判断A1单元格的值的类型是否为逻辑值。*/
function test() {
    console.log(WorksheetFunction.IsLogical(Range("A1").Value2))
}
```
