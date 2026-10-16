---
title: "Oct2Hex 方法"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/WorksheetFunction/%E6%96%B9%E6%B3%95/Oct2Hex%20%E6%96%B9%E6%B3%95.html
---
# Oct2Hex (方法)

将八进制数转换为十六进制数。

## 说明

- 如果参数 number 为负数，函数 OCT2HEX 将忽略 places，返回 10 位十六进制数。
- 如果参数不是有效的八进制数，函数 OCT2HEX 返回错误值 #NUM!。
- 如果函数 OCT2HEX 需要比 places 指定的更多的位数，将返回错误值 #NUM!。
- 如果 places 不是整数，则将被截尾取整。
- 如果 places 为非数值型，函数 OCT2HEX 返回错误值 #VALUE!。
- 如果 places 为负数，函数 OCT2HEX 返回错误值 #NUM!。

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| Arg1 | any | 必填 | Number - 要转换的八进制数。Number 不能多于 10 个八进制字符（30 个二进制位）。最高有效位为符号位，其余 29 位为数字位。负数用二进制数的补码表示。 |
| Arg2 | any | 可选 | Places - 要使用的字符数。如果省略 places，OCT2HEX 将用能表示它的最少字符来表示。当需要为返回的值填充前导 0（零）时，places 尤其有用。 |

## 返回值

String

## 示例

```javascript
/*本示例使用 Oct2Hex 方法分别将八进制数4567转换为十六进制数和八进制数1755指定4个字符数转换为十六进制数，并分别赋值给D2和D3单元格。*/
function test() {
    Range("D2").Value2 = WorksheetFunction.Oct2Hex(4567)
    Range("D3").Value2 = WorksheetFunction.Oct2Hex(1755, 4)
}
```

```javascript
/*本示例为C1单元格赋值，使用 Oct2Hex 方法将C1单元格中八进制数转换为十六进制数。*/
function test() {
    Range("C1").Value2 = 36
    console.log(WorksheetFunction.Oct2Hex(Range("C1").Value2))
}
```
