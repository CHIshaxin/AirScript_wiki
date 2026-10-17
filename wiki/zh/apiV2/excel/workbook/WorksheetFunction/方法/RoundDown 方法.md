---
title: "RoundDown (方法)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/WorksheetFunction/%E6%96%B9%E6%B3%95/RoundDown%20%E6%96%B9%E6%B3%95.html
updated: 1789700916000
hash: CDufIYj-
---

# RoundDown (方法)

将数字向零的方向向下舍入。

## 说明

- ROUNDDOWN 与 ROUND 功能相似，不同之处在于 ROUNDDOWN 总是向下舍入数字。
- 如果 num_digits 大于 0（零），则将 number 向下舍入到指定的小数位数。
- 如果 num_digits 为 0，则将 number 向下舍入到最接近的整数。
- 如果 num_digits 小于 0，则将 number 向下舍入到小数点左侧。

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| Arg1 | double | 必填 | Number - 要向下舍入的任何实数。 |
| Arg2 | double | 必填 | Num_digits - 要将数字舍入到的位数。 |

## 返回值

Double

## 示例

```javascript
/*本示例使用 RoundDown 方法将数字58.368和52236.12向下舍入到 0 (零) ，并分别赋值给B1和C1单元格。*/
function test() {
    Range("B1").Value2 = WorksheetFunction.RoundDown(58.368, 2)
    Range("C1").Value2 = WorksheetFunction.RoundDown(52236.12, -3)
}
```

```javascript
/*本示例为B1等单元格分别赋值，使用 RoundDown 方法将B1单元格中数字向下舍入到 0 (零)。*/
function test() {
    Range("B1").Value2 = -1.3315
    Range("B2").Value2 = 3
    console.log(WorksheetFunction.RoundDown(Range("B1").Value2, Range("B2").Value2))
}
```
