---
title: "Hex2Oct 方法"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/WorksheetFunction/%E6%96%B9%E6%B3%95/Hex2Oct%20%E6%96%B9%E6%B3%95.html
---
# Hex2Oct (方法)

将十六进制数转换为八进制数。

## 说明

- 如果 number 为负数，则 HEX2OCT 将忽略 places 并返回一个以 10 个字符表示的八进制数。
- 如果 number 为负数，则不能小于 FFE0000000；如果 number 为正数，则不能大于 1FFFFFFF。
- 如果 number 不是有效的十六进制数，则 HEX2OCT 将返回错误值 #NUM!。
- 如果 HEX2OCT 所需的字符数比 places 指定的字符数多，则将返回错误值 #NUM!。
- 如果 places 不是整数，则将被截尾取整。
- 如果 places 为非数值型，则 HEX2OCT 将返回错误值 #VALUE!。
- 如果 places 为负数，则 HEX2OCT 将返回错误值 #NUM!。

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| Arg1 | any | 必填 | Number - 要转换的十六进制数。Number 不能多于 10 个字符。最高位为符号位，其余 39 位为数字位。负数用二进制数的补码表示。 |
| Arg1 | any | 必填 | Places - 要使用的字符数。如果省略 places，HEX2OCT 将用能表示它的最少字符来表示。当需要为返回的值填充前导 0（零）时，places 尤其有用。 |

## 返回值

String

## 示例

```javascript
/*本示例使用 Hex2Oct 方法分别将十六进制数977转换为八进制数和十六进制数“03ED”指定使用8个字符转换为八进制数，并分别赋值给D2和D3单元格。*/
function test() {
    Range("D2").Value2 = WorksheetFunction.Hex2Oct(977)
    Range("D3").Value2 = WorksheetFunction.Hex2Oct("03ED", 8)
}
```

```javascript
/*本示例为A1单元格赋值，使用 Hex2Oct 方法将A1单元格中十六进制数转换为八进制数。*/
function test() {
    Range("A1").Value2 = "12AB"
    console.log(WorksheetFunction.Hex2Oct(Range("A1").Value2))
}
```
