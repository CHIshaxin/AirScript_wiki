---
title: "CountBlank (方法)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/WorksheetFunction/%E6%96%B9%E6%B3%95/CountBlank%20%E6%96%B9%E6%B3%95.html
updated: 1789700916000
hash: CMWvZw4g
---

# CountBlank (方法)

计算指定单元格区域中空单元格的个数。

## 说明

即使单元格中含有返回值为空文本 ("") 的公式，该单元格也会被计算在内，但包含零值的单元格不被计算在内。

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| Arg1 | Range | 必填 | 要计算其中的空白单元格个数的区域。 |

## 返回值

Double

## 示例

```javascript
/*本示例为A2等单元格分别赋值，并使用CountBlank方法计算指定单元格区域中空单元格的个数。*/
function test() {
    Range("A2").Value2 = 20
    Range("A3").Value2 = ""
    Range("A4").Value2 = "华为手机"
    Range("A5").Value2 = ""
    Range("A6").Value2 = 11.23
    let countblank1 = Application.WorksheetFunction.CountBlank(Range("A2:A6"))
    console.log(countblank1)
}
```

```javascript
/*本示例为B2等单元格分别赋值，使用CountBlank方法计算指定单元格区域中空单元格的个数，并将结果赋值于D3单元格。*/
function test() {
    Range("B2").Value2 = ""
    Range("B3").Value2 = 2
    Range("B4").Value2 = -13.56
    Range("B5").Value2 = ""
    Range("B6").Value2 = "地球"
    Range("B7").Value2 = "hello"
    Range("B8").Value2 = ""
    Range("D3").Value2 = Application.WorksheetFunction.CountBlank(Range("B2:B8"))
}
```
