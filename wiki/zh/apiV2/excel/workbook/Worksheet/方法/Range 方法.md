---
title: "Range (方法)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Worksheet/%E6%96%B9%E6%B3%95/Range%20%E6%96%B9%E6%B3%95.html
updated: 1789700916000
hash: BibZI_qg
---

# Range (方法)

返回一个 Range 对象，它代表一个单元格或单元格区域。

## 说明

如果在没有对象识别符时使用，则该属性是 `ActiveSheet.Range` 的快捷方式（它返回活动表的一个区域，如果活动表不是一张工作表，则该属性无效）。

当应用于 **[Range](./../../Range/Range 对象.html)** 对象时，该属性与 **[Range](./../../Range/Range 对象.html)** 对象相关。例如，如果选中单元格 C3，那么 `Selection.Range("B1")` 返回单元格 D3，因为它同 **Selection** 属性返回的 **[Range](./../../Range/Range 对象.html)** 对象相关。此外，代码 `ActiveSheet.Range("B1")` 总是返回单元格 B1。

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| Cell1 | any | 必填 | 区域名称。必须为采用宏语言的 A1 样式引用。可包括区域操作符（冒号）、相交区域操作符（空格）或合并区域操作符（逗号）。也可包括货币符号，但它们会被忽略掉。您可以在区域中任一部分使用局部定义名称。如果使用名称，则假定该名称使用的是宏语言。 |
| Cell2 | any | 可选 | 区域左上角和右下角的单元格。可以是一个包含单个单元格、整列或整行的 Range 对象，或者也可以是一个用宏语言为单个单元格命名的字符串。 |

## 示例

```javascript
/*此示例将 Sheet1 上 A1 单元格的值设置为 3.14159。*/
function test() {
    Application.Worksheets.Item("Sheet1").Range("A1").Value2 = 3.14159
}
```

```javascript
/*此示例在 Sheet1 的 A1 单元格中创建一个公式。*/
function test() {
    Application.Worksheets.Item("Sheet1").Range("A1").Formula = "=10*RAND()"
}
```

```javascript
/*此示例在 Sheet1 上的单元格区域 A1:D10 中进行循环。如果某个单元格的值小于 0.001，则此代码将用 0（零）来取代该值。*/
function test() {
    for (let i = 1; i <= Application.Worksheets.Item("Sheet1").Range("A1:D10").Count; i++) {
        if (Application.Worksheets.Item("Sheet1").Range("A1:D10").Item(i).Value2 < 0.001) {
            Application.Worksheets.Item("Sheet1").Range("A1:D10").Item(i).Value2 = 0
        }
    }
}
```

```javascript
/*此示例在名为“TestRange”的区域上进行循环，并显示该区域中空白单元格的个数。*/
function test() {
    let numBlanks = 0
    for (let i = 1; i <= Application.Range("TestRange").Count; i++) {
        if (Application.Range("TestRange").Item(i).Value2 == null) {
            numBlanks++
        }
    }
    console.log("There are " + numBlanks + " empty cells in this range")
}
```

```javascript
/*此示例将 Sheet1 中单元格区域 A1:C5 上的字体样式设置为斜体。此示例使用 Range 属性的语法 2。*/
function test() {
    Application.Worksheets.Item("Sheet1").Range(Application.Cells.Item(1, 1), Application.Cells.Item(5, 3)).Font.Italic = true
}
```
