---
title: "Ceiling (方法)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/WorksheetFunction/%E6%96%B9%E6%B3%95/Ceiling%20%E6%96%B9%E6%B3%95.html
updated: 1789700916000
hash: i8tu5INu
---

# Ceiling (方法)

返回向上舍入（远离零）到最接近的 significance 的倍数的 number。

## 说明

例如，如果要避免在价格中使用“分”，而产品价格为 ￥4.42, 那么可使用公式 `Ceiling(4.42,0.05)` 将价格进位到最近的“角”。

- 如果任一参数为非数值，**Ceiling** 将生成错误。
- 不论 number 的符号如何，向远离零的方向调整时，值都会向上舍入。如果 number 正好是 significance 的倍数，则无需进行任何舍入处理。
- 如果 number 和 significance 的符号不同，则 **Ceiling** 将生成错误。

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| Arg1 | double | 必填 | Number - 要舍入的值。 |
| Arg2 | double | 必填 | Significance - 用以进行舍入计算的倍数。 |

## 返回值

Double

## 示例

```javascript
/*本示例为B2等单元格分别赋值，并使用Ceiling方法将参数向上舍入为最接近的指定基数的倍数。*/
function test() {
    Range("B2").Value2 = 1
    Range("B3").Value2 = 2
    Range("B4").Value2 = 2.5
    Range("B5").Value2 = 4.55
    let ceiling1 = Application.WorksheetFunction.Ceiling(Range("B4").Value2, Range("B2").Value2)
    let ceiling2 = Application.WorksheetFunction.Ceiling(Range("B5").Value2, Range("B3").Value2)
    console.log(ceiling1)
    console.log(ceiling2)
}
```

```javascript
/*本示例使用Ceiling方法将参数向上舍入为最接近的指定基数的倍数，并将结果赋值于A4等单元格。*/
function test() {
    Range("A4").Value2 = Application.WorksheetFunction.Ceiling(-2.5, -2)
    Range("A5").Value2 = Application.WorksheetFunction.Ceiling(-1.5, 2)
    Range("A6").Value2 = Application.WorksheetFunction.Ceiling(1.5, 0.1)
    Range("A7").Value2 = Application.WorksheetFunction.Ceiling(0.234, 0.01)
}
```
