---
title: "Hex2Bin (方法)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/WorksheetFunction/%E6%96%B9%E6%B3%95/Hex2Bin%20%E6%96%B9%E6%B3%95.html
updated: 1789700916000
hash: 6VON0MBj
---

# Hex2Bin (方法)

将十六进制数转换为二进制数。

## 说明

- 如果 number 为负数，则 HEX2BIN 将忽略 places 并返回一个以 10 个字符表示的二进制数。
- 如果 number 为负数，则不能小于 FFFFFFFE00；如果 number 为正数，则不能大于 1FF。
- 如果 number 不是有效的十六进制数，则 HEX2BIN 将返回错误值 #NUM!。
- 如果 HEX2BIN 所需的字符数比 places 指定的字符数多，则将返回错误值 #NUM!。
- 如果 places 不是整数，则将被截尾取整。
- 如果 places 为非数值型，则 HEX2BIN 将返回错误值 #VALUE!。
- 如果 places 为负数，则 HEX2BIN 将返回错误值 #NUM!。

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| Arg1 | any | 必填 | Number - 要转换的十六进制数。Number 不能多于 10 个字符。最高位为符号位（从右算起第 40 位），其余 9 位是数字位。负数用二进制数的补码表示。 |
| Arg2 | any | 可选 | Places - 要使用的字符数。如果省略 places，HEX2BIN 将用能表示它的最少字符来表示。当需要为返回的值填充前导 0（零）时，places 尤其有用。 |

## 返回值

String

## 示例

```javascript
/*本示例使用 Hex2Bin 方法分别将十六进制数28转换为二进制数和十六进制数“BD”转换为二进制数，并分别赋值给A5和A6单元格。*/
function test() {
    Range("A5").Value2 = WorksheetFunction.Hex2Bin(28)
    Range("A6").Value2 = WorksheetFunction.Hex2Bin("BD")
}
```

```javascript
/*本示例为B1等单元格分别赋值，使用 Hex2Bin 方法将B1单元格中十六进制数转换为二进制数。*/
function test() {
    Range("B1").Value2 = "02D"
    Range("C1").Value2 = 8
    console.log(WorksheetFunction.Hex2Bin(Range("B1").Value2, Range("C1").Value2))
}
```
