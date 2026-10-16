---
title: "Floor 方法"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/WorksheetFunction/%E6%96%B9%E6%B3%95/Floor%20%E6%96%B9%E6%B3%95.html
---
# Floor (方法)

将 number 向下舍入（向零的方向），使其等于最接近的 significance 的倍数。

## 说明

- 如果两个参数中的任意一个为非数值型，则 FLOOR 将返回错误值 #VALUE!。
- 如果 number 和 significance 符号不同，则 FLOOR 将返回错误值 #NUM!。
- 不论 number 的符号如何，向远离零的方向调整时，值都会向下舍入。如果 number 正好是 significance 的倍数，则无需进行任何舍入处理。

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| Arg1 | double | 必填 | Number - 要进行舍入的数值。 |
| Arg2 | double | 必填 | Significance - 用以进行舍入计算的倍数。 |

## 返回值

Double

## 示例

```javascript
/*本示例使用 Floor 方法计算参数number向下舍入（向零的方向），使其等于最接近的参数significance的倍数，并分别赋值给B1等单元格。*/
function test() {
    Range("B1").Value2 = WorksheetFunction.Floor(5.6, 1)
    Range("B2").Value2 = WorksheetFunction.Floor(-3.8, -1)
}
```

```javascript
/*本示例为A1等单元格分别赋值，使用 Floor 方法计算A1单元格向下舍入（向零的方向），使其等于最接近的A2单元格的倍数。*/
function test() {
    Range("A1").Value2 = 0.08
    Range("A2").Value2 = 0.1
    console.log(WorksheetFunction.Floor(Range("A1").Value2, Range("A2").Value2))
}
```
