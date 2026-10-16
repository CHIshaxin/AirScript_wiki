---
title: "Odd 方法"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/WorksheetFunction/%E6%96%B9%E6%B3%95/Odd%20%E6%96%B9%E6%B3%95.html
---
# Odd (方法)

返回数字向上舍入到的最接近的奇数。

## 说明

- 如果 number 为非数字型，则 ODD 返回错误值 #VALUE!。
- 不论数字符号如何，都按远离零来进行舍入。如果 number 正好是奇数，则不进行任何舍入。

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| Arg1 | double | 必填 | Number - 要进行舍入的值。 |

## 返回值

Double

## 示例

```javascript
/*本示例使用 Odd 方法获取数字“0.123”和“3.456”向上舍入到的最接近的奇数，并分别赋值给A1和A2单元格。*/
function test() {
    Range("A1").Value2 = WorksheetFunction.Odd(0.123)
    Range("A2").Value2 = WorksheetFunction.Odd(3.456)
}
```

```javascript
/*本示例为B1单元格赋值，使用 Odd 方法获取B1单元格中数字向上舍入到的最接近的奇数。*/
function test() {
    Range("B1").Value2 = -6
    console.log(WorksheetFunction.Odd(Range("B1").Value2))
}
```
