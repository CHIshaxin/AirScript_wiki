---
title: "TextToColumns (方法)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Range/%E6%96%B9%E6%B3%95/TextToColumns%20%E6%96%B9%E6%B3%95.html
updated: 1789700916000
hash: BGA6cVUV
---

# TextToColumns (方法)

将包含文本的一列单元格分解为若干列。

## 说明

下表显示了使用不同的导入设置向 ET 中导入文本时的结果。数字结果显示在最右边的列中。

| 系统小数分隔符 | 系统千位分隔符 | 小数分隔符值 | 千位分隔符值 | 原始文本 | 单元格的值（数据类型） |
| --- | --- | --- | --- | --- | --- |
| 句号 | 逗号 | 逗号 | 句号 | 123.123,45 | 123,123.45（数字） |
| 句号 | 逗号 | 逗号 | 逗号 | 123.123,45 | 123.123,45（文本） |
| 逗号 | 句号 | 逗号 | 句号 | 123,123.45 | 123,123.45（数字） |
| 句号 | 逗号 | 句号 | 逗号 | 123 123.45 | 123 123.45（文本） |
| 句号 | 逗号 | 句号 | 空格 | 123 123.45 | 123,123.45（数字） |

| **[XlColumnDataType](./../../枚举/XlColumnDataType 枚举.html)** 可为以下 **[XlColumnDataType](./../../枚举/XlColumnDataType 枚举.html)** 常量之一。 |
| --- |
| **xlGeneralFormat**：表示常规 |
| **xlTextFormat**：表示文本 |
| **xlMDYFormat**：表示 MDY 日期 |
| **xlDMYFormat**：表示 DMY 日期 |
| **xlYMDFormat**：表示 YMD 日期 |
| **xlMYDFormat**：表示 MYD 日期 |
| **xlDYMFormat**：表示 DYM 日期 |
| **xlYDMFormat**：表示 YDM 日期 |
| **xlEMDFormat**：表示 EMD 日期 |
| **xlSkipColumn**：表示跳过列 |

只有在安装并选定了中国台湾地区语言时才可以使用 **xlEMDFormat**。**xlEMDFormat** 常量指定使用中国台湾地区纪元日期。

列说明符可为任意顺序。在输入数据中，如果某一特定列没有给定列说明符，则使用 **xlGeneralFormat** 设置对该列进行分列。该示例跳过了源数据的第三列，第一列作为文本进行分列，而其余各列均使用 **xlGeneralFormat** 设置进行分列。

`Array(Array(3, 9), Array(1, 2))`

如果源数据有固定宽度的列，则每个两元素数组的第一个元素将指定列中起始字符的位置（以整数的形式，0（零）表示第一个字符）。第二个元素以从 1 到 9 的数字指定该列的拆分选项，如上表所示。

下例从固定宽度的文件中拆分出两列，第一列从行起始处开始，长度为 10 个字符。第二列从第 15 个字符开始，直至行尾。为避免包含从第 10 个字符到第 15 个字符之间的字符，ET 加入了一个被跳过的列数据项。

`Array(Array(0, 1), Array(10, 9), Array(15, 1))`

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| Destination | any | 可选 | 一个 **[Range](./../../Range/Range 对象.html)** 对象，指定 ET 放置结果的位置。如果该区域大于一个单元格，则使用左上角的单元格。 |
| DataType | Default | 可选 | 将被拆分到多列中的文本的格式。 |
| TextQualifier | Default | 可选 | 指定是将单引号、双引号用作文本分隔符还是不使用引号。 |
| ConsecutiveDelimiter | any | 可选 | 如果为 **True**，则 ET 将连续分隔符视为一个分隔符。默认值为 **False**。 |
| Tab | any | 可选 | 如果为 **True**，则 *DataType* 为 **xlDelimited** 并将制表符作为分隔符。默认值为 **False**。 |
| Semicolon | any | 可选 | 如果为 **True**，则 *DataType* 为 **xlDelimited** 并将分号作为分隔符。默认值为 **False**。 |
| Comma | any | 可选 | 如果为 **True**，则 *DataType* 为 **xlDelimited** 并将逗号作为分隔符。默认值为 **False**。 |
| Space | any | 可选 | 如果为 **True**，则 *DataType* 为 **xlDelimited** 并将空格字符作为分隔符。默认值为 **False**。 |
| Other | any | 可选 | 如果为 **True**，则 *DataType* 为 **xlDelimited** 并将 *OtherChar* 参数指定的字符作为分隔符。默认值为 **False**。 |
| OtherChar | any | 可选 | （如果 *Other* 为 **True**，则为必选项）。当 *Other* 为 **True** 时的分隔符。如果指定了多个字符，则仅使用字符串中的第一个字符而忽略剩余字符。 |
| FieldInfo | any | 可选 | 包含单列数据相关分列信息的数组。对该参数的解释取决于 *DataType* 的值。如果此数据由分隔符分隔，则该参数为由两元素数组组成的数组，其中每个两元素数组指定一个特定列的转换选项。第一个元素为列标（从 1 开始），第二个元素是 xlColumnDataType 的常量之一，用于指定分列方式。 |
| DecimalSeparator | any | 可选 | 识别数字时，ET 使用的小数分隔符。默认设置为系统设置。 |
| ThousandsSeparator | any | 可选 | 识别数字时，ET 使用的千位分隔符。默认设置为系统设置。 |
| TrailingMinusNumbers | any | 可选 | 以减号字符开始的数字。 |

## 返回值

Variant

## 示例

```javascript
/*本示例将剪贴板中的内容（以空格分隔的文字表格）转换为 Sheet1 中单独的列内。可用“记事本”或“写字板”（或其他文本编辑器）创建简单的以空格分隔的表格，将该表格复制到剪贴板，再切换到 ET，然后运行本示例。*/
function test() {
    Worksheets.Item("Sheet1").Activate()
    ActiveSheet.Paste()
    Selection.TextToColumns(null, xlDelimited, null, true, null, null, null, true)
}
```

```javascript
/*本示例将 A1 单元格的值设置为“test1,test2,test3”，将 A1 单元格的值分解为若干列，以逗号为分隔符，不使用引号作为文本分隔符，多个分隔符视为一个，填入 E1:G1 单元格区域中。*/
function test() {
    Range("A1").Value2 = "test1,test2,test3"
    Range("A1").TextToColumns(Range("E1:G1"), xlDelimited, xlTextQualifierNone, true, false, false, true)
}
```
