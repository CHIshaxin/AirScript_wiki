---
title: "Item (方法)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Range/%E6%96%B9%E6%B3%95/Item%20%E6%96%B9%E6%B3%95.html
updated: 1789700916000
hash: Bv2bap4g
---

# Item (方法)

返回一个 **[Range](./../../Range/Range 对象.html)** 对象，它代表对指定区域某一偏移量处的区域。

## 说明

语法 1 使用行号和列号或列标作为索引参数。关于此语法的详细信息，请参阅 **[Range](./../../Range/Range 对象.html)** 对象。**RowIndex** 和 **ColumnIndex** 参数为相对偏移量。换句话说，如果 **RowIndex** 指定为 1，则返回区域内第一行中的单元格，而非工作表的第一行。例如，如果选定区域为单元格 C3，则 Selection.Cells(2, 2) 返回单元格 D4（使用 **Item** 属性可在原始区域之外进行索引）。

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| RowIndex | any | 必填 | 要访问的单元格的索引号，顺序为从左到右，然后往下。 Range.Item(1) 返回区域左上角单元格； Range.Item(2) 返回紧靠左上角单元格右侧的单元格。 |
| ColumnIndex | any | 可选 | 指明要访问的单元格所在列的列号的数字或字符串，1 或 “A”表示区域中的第一列。 |

## 示例

```javascript
/*本示例基于单元格 A1 的内容填写 Sheet1 的单元格区域 A2。*/
function test() {
    Application.Worksheets.Item("Sheet1").Range("A1").Item(2).FillDown()
}
```

```javascript
/*本示例将 B 列第二个单元格的值设置为 b2*/
function test() {
    Range("B:B").Item(2).Value2 = 'b2'
}
```
