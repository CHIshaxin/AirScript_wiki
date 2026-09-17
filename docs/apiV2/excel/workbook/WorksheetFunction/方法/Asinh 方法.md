---
title: "Asinh 方法"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/WorksheetFunction/%E6%96%B9%E6%B3%95/Asinh%20%E6%96%B9%E6%B3%95.html
---
# Asinh (方法)

返回数字的反双曲正弦值。反双曲正弦值的双曲正弦值为 *Arg1*，因此 Asinh(Sinh(number)) 等于 *Arg1*。

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| Arg1 | double | 必填 | 任意实数。 |

## 返回值

Double

## 示例

```javascript
/*本示例使用Asinh计算反双曲正弦值，并将结果分别赋值A1等单元格*/
function test() {
    Range("A1").Value2 = Application.WorksheetFunction.Asinh(4)
    Range("A2").Value2 = Application.WorksheetFunction.Asinh(16.66)
    Range("A3").Value2 = Application.WorksheetFunction.Asinh(-20)
    Range("A4").Value2 = Application.WorksheetFunction.Asinh(48)
}
```

```javascript
/*本示例为C3等单元格分别赋值，并使用Asinh计算反双曲正弦值*/
function test() {
    Range("C3").Value2 = 3
    Range("D3").Value2 = 8
    Range("E3").Value2 = -41.23
    Range("F3").Value2 = 0
    let asinh1 = Application.WorksheetFunction.Asinh(Range("C3").Value2)
    let asinh2 = Application.WorksheetFunction.Asinh(Range("D3").Value2)
    let asinh3 = Application.WorksheetFunction.Asinh(Range("E3").Value2)
    let asinh4 = Application.WorksheetFunction.Asinh(Range("F3").Value2)
    console.log(asinh1)
    console.log(asinh2)
    console.log(asinh3)
    console.log(asinh4)
}
```
