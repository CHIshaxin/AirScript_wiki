---
title: "Bin2Dec (方法)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/WorksheetFunction/%E6%96%B9%E6%B3%95/Bin2Dec%20%E6%96%B9%E6%B3%95.html
updated: 1789700916000
hash: jCuCnlB0
---

# Bin2Dec (方法)

将二进制数转换为十进制数。

## 说明

如果 number 不是有效的二进制数或多于 10 个字符（10 位），则 Bin2Dec 将生成一个错误值。

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| Arg1 | any | 必填 | 要转换的二进制数。Number 不能多于 10 个字符（10 位）。最高位为符号位，其余 9 位为数字位。负数用二进制数的补码表示。 |

## 返回值

String

## 示例

```javascript
/*本示例使用Bin2Dec方法将二进制数转换为十进制数，并将结果赋值于A2等单元格。*/
function test() {
    Range("A2").Value2 = Application.WorksheetFunction.Bin2Dec(1100100)
    Range("A3").Value2 = Application.WorksheetFunction.Bin2Dec(11101)
    Range("A4").Value2 = Application.WorksheetFunction.Bin2Dec(1001010)
    Range("A5").Value2 = Application.WorksheetFunction.Bin2Dec(10)
}
```

```javascript
/*本示例为C3等单元格分别赋值，并使用Bin2Dec方法将二进制数转换为十进制数。*/
function test() {
    Range("C3").Value2 = 1100
    Range("C4").Value2 = 101101
    Range("C5").Value2 = 1001111
    let bin2dec1 = Application.WorksheetFunction.Bin2Dec(Range("C3").Value2)
    let bin2dec2 = Application.WorksheetFunction.Bin2Dec(Range("C4").Value2)
    let bin2dec3 = Application.WorksheetFunction.Bin2Dec(Range("C5").Value2)
    console.log(bin2dec1)
    console.log(bin2dec2)
    console.log(bin2dec3)
}
```
