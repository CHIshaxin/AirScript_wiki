---
title: "Round 方法"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/WorksheetFunction/%E6%96%B9%E6%B3%95/Round%20%E6%96%B9%E6%B3%95.html
---
# Round (方法)

将数字舍入到指定位数。

## 说明

- 如果 num_digits 大于 0（零），则会将 number 舍入到指定的小数位数。
- 如果 num_digits 为 0，则会将 number 舍入到最接近的整数。
- 如果 num_digits 小于 0，则会将 number 舍入到小数点左侧。

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| Arg1 | double | 必填 | Number - 要进行舍入的数字。 |
| Arg2 | double | 必填 | Num_digits - 指定要将数字舍入到的位数。 |

## 返回值

Double

## 示例

```javascript
/*本示例使用 Round 方法将数字123.654和-235.654分别舍入到2位和-1位，并分别赋值给B1和C1单元格。*/
function test() {
    Range("B1").Value2 = WorksheetFunction.RoundDown(123.654, 2)
    Range("C1").Value2 = WorksheetFunction.RoundDown(-235.654, -1)
}
```

```javascript
/*本示例为A1等单元格分别赋值，使用 Round 方法将A1单元格数字舍入到指定位数。*/
function test() {
    Range("A1").Value2 = -100.3615
    Range("B1").Value2 = 2
    console.log(WorksheetFunction.Round(Range("A1").Value2, Range("B1").Value2))
}
```
