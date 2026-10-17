---
title: "CountIf (方法)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/WorksheetFunction/%E6%96%B9%E6%B3%95/CountIf%20%E6%96%B9%E6%B3%95.html
updated: 1789700916000
hash: CKGSta0W
---

# CountIf (方法)

计算区域中满足给定条件的单元格的个数。

## 说明

- 可以在条件中使用通配符，包括问号 (?) 和星号 (*)。问号可匹配任意的单个字符；星号可匹配任意一串字符。如果要查找实际的问号或星号，则请在该字符前键入一个波形符 (~)。

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| Arg1 | Range | 必填 | 要计算其中满足条件的单元格个数的单元格区域。 |
| Arg2 | any | 必填 | 用于定义哪些单元格将被计算在内的条件，其形式可以为数字、表达式、单元格引用或文本。例如，条件可以表示为 32、"32"、">32"、"apples" 或 B4。 |

## 返回值

Double

## 示例

```javascript
/*本示例为B3等单元格分别赋值，并使用CountIf方法计算区域中满足给定条件的单元格的个数。*/
function test() {
    Range("B3").Value2 = "苹果"
    Range("B4").Value2 = "西瓜"
    Range("B5").Value2 = "猕猴桃"
    Range("B6").Value2 = "西瓜"
    Range("B7").Value2 = "桃子"
    let countif1 = Application.WorksheetFunction.CountIf(Range("B2:B7"), "西瓜")
    console.log(countif1)
}
```

```javascript
/*本示例为E2等单元格分别赋值，并使用CountIf方法计算区域中满足给定条件的单元格的个数，并将结果赋值于C3单元格。*/
function test() {
    Range("E2").Value2 = 20
    Range("E3").Value2 = 20.1
    Range("E4").Value2 = 15
    Range("E5").Value2 = 40
    Range("E6").Value2 = 78
    Range("C3").Value2 = Application.WorksheetFunction.CountIf(Range("E2:E6"), ">20")
}
```
