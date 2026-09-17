---
title: "CountA 方法"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/WorksheetFunction/%E6%96%B9%E6%B3%95/CountA%20%E6%96%B9%E6%B3%95.html
---
# CountA (方法)

计算非空单元格及参数列表中值的个数。

## 说明

使用 CountA 可以计算区域或数组中包含数据的单元格的个数。

- 值可以是任何信息类型，包括错误值和空文本 ("")，但不包括空单元格。
- 如果参数为数组或引用，则只会使用该数组或引用中的值。数组或引用中的空单元格和文本值将被忽略。
- 如果不需要统计逻辑值、文本或错误值，则请使用 Count 函数。

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| Arg1 | any | 必填 | Arg30 代表要计算在内的值的 1 到 30 个参数。 |

## 返回值

Double

## 示例

```javascript
/*本示例为B2等单元格分别赋值，并使用CountA方法计算非空单元格中值的个数。*/
function test() {
    Range("B2").Value2 = 5
    Range("B3").Value2 = 3
    Range("B4").Value2 = "西瓜"
    Range("B5").Value2 = ""
    Range("B6").Value2 = 1.5
    let counta1 = Application.WorksheetFunction.CountA(Range("B2:B6"))
    console.log(counta1)
}
```

```javascript
/*本示例使用CountA方法计算参数列表中值的个数，并将结果赋值于D3等单元格。*/
function test() {
    Range("D3").Value2 = Application.WorksheetFunction.CountA(2, 4, "香蕉", "苹果", true)
    Range("D4").Value2 = Application.WorksheetFunction.CountA(12, 13, 14, 15, 16)
    Range("D5").Value2 = Application.WorksheetFunction.CountA("篮球", "乒乓球", false)
}
```
