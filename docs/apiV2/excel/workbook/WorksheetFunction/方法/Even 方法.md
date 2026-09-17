---
title: "Even 方法"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/WorksheetFunction/%E6%96%B9%E6%B3%95/Even%20%E6%96%B9%E6%B3%95.html
---
# Even (方法)

返回向上舍入到最接近的偶数的数字。可以使用此函数处理成对出现的项目。例如，一个包装箱一行可容纳一宗或两宗货物。只有当这些货物的数目向上舍入到最接近的偶数，与包装箱的容量相匹配时，包装箱才会装满。

## 说明

- 如果 number 为非数值型，则 EVEN 将返回错误值 #VALUE!。
- 不论 number 的符号如何，向远离零的方向调整时，值都会向上舍入。如果 number 是偶数，则不会进行任何舍入处理。

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| Arg1 | double | 必填 | Number - 是要进行舍入的值。 |

## 返回值

Double

## 示例

```javascript
/*本示例分别赋值A1等单元格，使用 Even 方法计算向上舍入到最接近的偶数的数字。*/
function test() {
    Range("A1").Value2 = 6
    Range("A2").Value2 = 3
    console.log(WorksheetFunction.Even(Range("A1").Value2))
    console.log(WorksheetFunction.Even(Range("A2").Value2))
}
```

```javascript
/*本示例使用 Even 方法计算向上舍入到最接近的偶数的数字，并分别赋值给D1等单元格。*/
function test() {
    Range("D1").Value2 = WorksheetFunction.Even(0)
    Range("D2").Value2 = WorksheetFunction.Even(-3.3)
}
```
