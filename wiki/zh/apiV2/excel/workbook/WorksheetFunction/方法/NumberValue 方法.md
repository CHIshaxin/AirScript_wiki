---
title: "NumberValue (方法)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/WorksheetFunction/%E6%96%B9%E6%B3%95/NumberValue%20%E6%96%B9%E6%B3%95.html
updated: 1789700916000
hash: DQ2rvMxJ
---

# NumberValue (方法)

以独立于区域设置的方式将文本转换为数字。

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| Arg1 | string | 必填 | 表示要转换的数字的字符串。 |
| Arg2 | string | 必填 | 用作字符串中小数分隔符的字符。 |
| Arg3 | string | 必填 | 用作字符串中的组分隔符的字符。 |

## 返回值

Double

## 示例

```javascript
/*本示例使用 NumberValue 方法以独立于区域设置的方式将文本转换为数字，并分别赋值给B2和B3单元格。*/
function test() {
    Range("B2").Value2 = WorksheetFunction.NumberValue("12.45,69", ",", ".")
    Range("B3").Value2 = WorksheetFunction.NumberValue("12,34%", ",", ";")
}
```

```javascript
/*本示例为A1等单元格分别赋值，使用 NumberValue 方法以独立于区域设置的方式将文本转换为数字。*/
function test() {
    Range("A1").Value2 = "23,45.56"
    Range("A2").Value2 = "."
    Range("A3").Value2 = ","
    console.log(WorksheetFunction.NumberValue(Range("A1").Value2, Range("A2").Value2, Range("A3").Value2))
}
```
