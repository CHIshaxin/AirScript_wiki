---
title: "Count 方法"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/WorksheetFunction/%E6%96%B9%E6%B3%95/Count%20%E6%96%B9%E6%B3%95.html
---
# Count (方法)

计算包含数字的单元格的个数，以及参数列表中的数字的个数。

## 说明

使用 Count 可以获得区域或数字数组中数字字段的输入项个数。

- 作为数字、日期或用文本表示的数字的参数会被计算在内。
- 直接键入参数列表的逻辑值和数字的文本表示也包括在内。
- 如果参数为错误值或无法转换为数字的文本，则将被忽略。
- 如果参数为数组或引用，则只有该数组或引用中的数字将被计算在内。数组或引用中的空单元格、逻辑值、文本或错误值将被忽略。
- 如果想统计逻辑值、文本或错误值，则请使用 CountA 函数。

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| Arg1 | any | 必填 | Arg30 可以包含或引用各种不同类型的数据的 1 到 30 个参数，但只有数字才计算在内。 |

## 返回值

Double

## 示例

```javascript
/*本示例为C2等单元格分别赋值，并使用Count方法计算包含数字的单元格的个数。*/
function test() {
    Range("C2").Value2 = 1
    Range("C3").Value2 = 2
    Range("C4").Value2 = "西瓜"
    Range("C5").Value2 = 12
    Range("C6").Value2 = 13.5
    let count1 = Application.WorksheetFunction.Count(Range("C2:C6"))
    console.log(count1)
}
```

```javascript
/*本示例使用Count方法计算包含数字的单元格的个数，并将结果赋值于F3等单元格。*/
function test() {
    Range("F3").Value2 = Application.WorksheetFunction.Count(12, 24, 35, "香蕉", "苹果", 30)
    Range("F4").Value2 = Application.WorksheetFunction.Count(12.3, 12.5, 15.33)
    Range("F5").Value2 = Application.WorksheetFunction.Count("足球", "篮球", "乒乓球")
}
```
