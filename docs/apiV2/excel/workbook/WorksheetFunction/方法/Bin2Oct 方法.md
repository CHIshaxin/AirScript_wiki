---
title: "Bin2Oct 方法"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/WorksheetFunction/%E6%96%B9%E6%B3%95/Bin2Oct%20%E6%96%B9%E6%B3%95.html
---
# Bin2Oct (方法)

将二进制数转换为八进制数。

## 说明

- 如果 number 不是有效的二进制数或多于 10 个字符（10 位），则 Bin2Oct 将生成一个错误。
- 如果 number 为负数，则 Bin2Oct 将忽略 places，并返回一个以 10 个字符表示的八进制数。
- 如果 Bin2Oct 所需的字符数比 places 指定的字符数多，则将生成一个错误。
- 如果 places 不是整数，则将被截尾取整。
- 如果 places 为非数值型，则 Bin2Oct 将生成一个错误。
- 如果 places 为负数，则 Bin2Oct 将生成一个错误。

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| Arg1 | any | 必填 | 要转换的二进制数。Number 不能多于 10 个字符（10 位）。最高位为符号位，其余 9 位为数字位。负数用二进制数的补码表示。 |
| Arg2 | any | 可选 | 要使用的字符数。如果省略 places，Bin2Oct 将用能表示此数的最少字符来表示。当需要为返回的值填充前导 0（零）时，places 尤其有用。 |

## 返回值

String

## 示例

```javascript
/*本示例使用Bin2Oct方法将二进制数转换为八进制数，并将结果赋值于C2等单元格。*/
function test() {
    Range("C2").Value2 = Application.WorksheetFunction.Bin2Oct(10101010)
    Range("C3").Value2 = Application.WorksheetFunction.Bin2Oct(1010110, 4)
    Range("C4").Value2 = Application.WorksheetFunction.Bin2Oct(10000)
}
```

```javascript
/*本示例为A2等单元格分别赋值，并使用Bin2Oct方法将二进制数转换为八进制数。*/
function test() {
    Range("A2").Value2 = 1001
    Range("A3").Value2 = 1110
    Range("A4").Value2 = 1111111111
    let bin2hex1 = Application.WorksheetFunction.Bin2Oct(Range("A2").Value2, 3)
    let bin2hex2 = Application.WorksheetFunction.Bin2Oct(Range("A3").Value2)
    let bin2hex3 = Application.WorksheetFunction.Bin2Oct(Range("A4").Value2)
    console.log(bin2hex1)
    console.log(bin2hex2)
    console.log(bin2hex3)
}
```
