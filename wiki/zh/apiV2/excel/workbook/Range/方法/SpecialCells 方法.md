---
title: "SpecialCells (方法)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Range/%E6%96%B9%E6%B3%95/SpecialCells%20%E6%96%B9%E6%B3%95.html
updated: 1789700916000
hash: BQrIdYa3
---

# SpecialCells (方法)

返回一个 **[Range](./../../Range/Range 对象.html)** 对象，该对象代表与指定类型和值匹配的所有单元格。

## 说明

使用 XlSpecialCellsValue 枚举指定要包含在结果中的具有特定类型的值的单元格。

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| Type | XlCellType | 必填 | 要包含的单元格。 |
| Value | any | 可选 | 如果 *Type* 为 **xlCellTypeConstants** 或 **xlCellTypeFormulas**，则该参数可用于确定结果中应包含哪几类单元格。将这些值相加可使此方法返回多种类型的单元格。默认情况下，将选择所有常量或公式，无论类型如何。 |

## 返回值

Range

## 示例

```javascript
/*本示例选定工作表 Sheet1 中已用区域的最后一个单元格。*/
function test() {
    Worksheets.Item("Sheet1").Activate()
    ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Activate()
}
```

```javascript
/*本示例给 C3 单元格添加批注 test，并显示 A1:C5 单元格区域含有批注的单元格地址*/
function test() {
    Range("C3").AddComment("test")
    console.log(Range("A1:C5").SpecialCells(xlCellTypeComments).Address())
}
```
