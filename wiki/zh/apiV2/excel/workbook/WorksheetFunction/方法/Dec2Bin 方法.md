---
title: "Dec2Bin (方法)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/WorksheetFunction/%E6%96%B9%E6%B3%95/Dec2Bin%20%E6%96%B9%E6%B3%95.html
updated: 1789700916000
hash: C0pdoWOm
---

# Dec2Bin (方法)

将十进制数转换为二进制数。

## 说明

- 如果 number < -512 或 number > 511，则 DEC2BIN 将返回错误值 #NUM!。
- 如果 number 为非数值型，则 DEC2BIN 将返回错误值 #VALUE!。
- 如果 DEC2BIN 所需的字符数多于 places 指定的字符数，则将返回错误值 #NUM!。
- 如果 places 不是整数，则将被截尾取整。
- 如果 places 为非数值型，则 DEC2BIN 将返回错误值 #VALUE!。
- 如果 places 为零或为负数，则 DEC2BIN 将返回错误值 #NUM!。

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| Arg1 | any | 必填 | Number - 要转换的十进制整数。如果 number 是负数，则将忽略有效位置值，而且 DEC2BIN 将返回一个以 10 个字符（10 位）表示的二进制数，最高位为符号位，其余 9 位是数字位。负数用二进制数的补码表示。 |
| Arg2 | any | 可选 | Places - 要使用的字符数。如果省略 places，DEC2BIN 将用能表示它的最少字符来表示。当需要为返回的值填充前导 0（零）时，places 尤其有用。 |

## 返回值

String

## 示例

```javascript
/*本示例使用Dec2Bin方法将十进制数转换为二进制数，并将结果赋值于C2等单元格。*/
function test() {
    Range("C2").Value2 = Application.WorksheetFunction.Dec2Bin(4, 4)
    Range("C3").Value2 = Application.WorksheetFunction.Dec2Bin(123)
    Range("C4").Value2 = Application.WorksheetFunction.Dec2Bin(78)
}
```

```javascript
/*本示例为A3等单元格分别赋值，并使用Dec2Bin方法将十进制数转换为二进制数。*/
function test() {
    Range("A3").Value2 = 9
    Range("A4").Value2 = -100
    Range("A5").Value2 = 500.5
    let dec2bin1 = Application.WorksheetFunction.Dec2Bin(Range("A3").Value2, 4)
    let dec2bin2 = Application.WorksheetFunction.Dec2Bin(Range("A4").Value2)
    let dec2bin3 = Application.WorksheetFunction.Dec2Bin(Range("A5").Value2)
    console.log(dec2bin1)
    console.log(dec2bin2)
    console.log(dec2bin3)
}
```
