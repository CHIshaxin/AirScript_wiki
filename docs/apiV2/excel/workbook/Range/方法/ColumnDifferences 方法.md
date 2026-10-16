---
title: "ColumnDifferences 方法"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Range/%E6%96%B9%E6%B3%95/ColumnDifferences%20%E6%96%B9%E6%B3%95.html
---
# ColumnDifferences (方法)

返回一个 **[Range](./../../Range/Range 对象.html)** 对象，该对象表示每列中所有与比较单元格内容不同的单元格。

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| Comparison | any | 必填 | 用于对指定区域进行比较的单个单元格。 |

## 返回值

Range

## 示例

```javascript
/*本示例选定工作表 Sheet1 上 A 列中与单元格 A4 内容不同的单元格。*/
function test() {
    //Setting up data to be compared
    Range("A1").Value(undefined, "Rod")
    Range("A2").Value(undefined, "Bill")
    Range("A3").Value(undefined, "John")
    Range("A4").Value(undefined, "Rod")
    Range("A5").Value(undefined, "Kelly")
    Range("A6").Value(undefined, "Rod")
    Range("A7").Value(undefined, "Paddy")
    Range("A8").Value(undefined, "Rod")
    Range("A9").Value(undefined, "Rod")
    Range("A10").Value(undefined, "Rod")
    //Code to do the comparison, selects the values that are unlike A1
    Worksheets.Item("Sheet1").Activate()
    let r1 = ActiveSheet.Columns.Item("A").ColumnDifferences(ActiveSheet.Range("A1"))
    r1.Select()
}
```

```javascript
/*本示例将数字1至5填入 A1:A5 单元格，A6 单元格的值设置为 2，最后选中 A 列与 A6 单元格内容不同的单元格*/
function test() {
    Range("A1").Value2 = 1
    Range("A1").AutoFill(Range("A1:A5"), xlFillValues)
    Range("A6").Value2 = 2
    let r1 = Columns.Item("A").ColumnDifferences(Range("A6"))
    r1.Select()
}
```
