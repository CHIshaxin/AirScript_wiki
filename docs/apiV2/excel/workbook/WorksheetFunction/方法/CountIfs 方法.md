---
title: "CountIfs 方法"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/WorksheetFunction/%E6%96%B9%E6%B3%95/CountIfs%20%E6%96%B9%E6%B3%95.html
---
# CountIfs (方法)

计算区域中满足多个条件的单元格的个数。

## 说明

- 仅在每个单元格中指定的对应条件都为 True 时，才会将区域内的每个单元格都计算在内。
- 如果任何参数中的单元格为空单元格，CountIfs 都会将其作为 0 值处理。
- 可以在条件中使用通配符，包括问号 (?) 和星号 (\*)。问号可匹配任意的单个字符；星号可匹配任意一串字符。如果要查找实际的问号或星号，则请在该字符前键入一个波形符 (~)。

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| Arg1 | Range | 必填 | 在其中计算相关条件的一个或多个区域。 |
| Arg2 | any | 必填 | Arg30 用于定义哪些单元格将被计算在内的一个或多个条件，其形式可以为数字、表达式、单元格引用或文本。例如，条件可以表示为 32、"32"、">32"、"apples" 或 B4。 |

## 返回值

Double

## 示例

```javascript
/*本示例为D2等单元格分别赋值，使用CountIfs方法计算区域中满足多个条件的单元格的个数，并将结果赋值于F3单元格。*/
function test() {
    Range("D2").Value2 = 60
    Range("D3").Value2 = 60
    Range("D4").Value2 = 50
    Range("D5").Value2 = 50
    Range("D6").Value2 = 50
    Range("E2").Value2 = "高三六班"
    Range("E3").Value2 = "高三一班"
    Range("E4").Value2 = "高三六班"
    Range("E5").Value2 = "高三六班"
    Range("E6").Value2 = "高三七班"
    Range("F3").Value2 = Application.WorksheetFunction.CountIfs(Range("E2:E6"), "高三六班", Range("D2:D6"), 50)
}
```

```javascript
/*本示例为B3等单元格分别赋值，并使用CountIfs方法计算区域中满足多个条件的单元格的个数。*/
function test() {
    Range("B3").Value2 = "苹果"
    Range("B4").Value2 = "桃子"
    Range("B5").Value2 = "猕猴桃"
    Range("B6").Value2 = "桃子"
    Range("C3").Value2 = 30
    Range("C4").Value2 = 40
    Range("C5").Value2 = 35
    Range("C6").Value2 = 50
    let countifs1 = Application.WorksheetFunction.CountIfs(Range("B3:B6"), "桃子", Range("C3:C6"), ">30")
    console.log(countifs1)
}
```
