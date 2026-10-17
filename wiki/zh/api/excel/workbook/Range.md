---
title: "Range"
source: https://airsheet.wps.cn/docs/api/excel/workbook/Range.html
updated: 1789700916000
hash: DpwZtevl
---

# Range

表示一个单元格、一行、一列、一个包含单个或若干连续单元格区域的选定单元格范围

Range 对象是工作表中部分单元格的集合，可以对工作表中某一区域的单元格进行操作

下方是 Range 对象的使用示例，Range 对象的具体属性和方法请参阅下方的列表。

```js
// 选择A1:A10单元格
let range = Application.Range('A1:A10')
// 把A1:A10的值设为对应行号
for (let i = 1; i <= 10; i++) {
  range.Item(i).Value = i.toString()
}
// 打印
printText(range) // 1...10
// 给单元格加上前缀和后缀
range.Each(cell => {
  cell.Value = 'prefix_' + cell.Text + '_suffix'
})
printText(range) //prefix_1_suffix...prefix_10_suffix

function printText(r) {
  r.Each(cell => {
    console.log(cell.Text)
  })
}
```

#### 属性列表

| 属性名 | 数据类型 | 简介 |
| --- | --- | --- |
| [Count](#count) | Number | 区域中单元格的数量 |
| [Text](#text) | String | 【只读】读取单元格格式化文本 |
| [Value/Value2](#value) | any/[][]any | 读写单元格中的值 |
| [FormatConditions](#formatConditions) | [FormatConditions](FormatConditions.md) | 用于控制 Excel 中的条件格式 |
| [Formula](#formula) | String | 以 A1 样式表示法表示的对象的隐式交叉的公式 |
| [FormulaArray](#formulaarray) | String | 返回或设置区域的数组公式 |
| [NumberFormat](#numberformat) | String | 获取或者设置区域的数字格式 |
| [Hidden](#hidden) | Boolean | 行或者列的隐藏 |
| [Interior.Color](#interior-color) | String | 内部颜色的十六进制 RGB |
| [HorizontalAlignment](#horizontalalignment) | [Enum.XlHAlign](Enum.md#xlhalign) | 设置区域的水平对齐方式 |
| [VerticalAlignment](#verticalalignment) | [Enum.XlVAlign](Enum.md#xlvalign) | 设置区域的垂直对齐方式 |
| [WrapText](#wraptext) | Boolean | 获取或者设置区域自动换行 |
| [IndentLevel](#indentlevel) | Number | 单元格缩进 |
| [MergeArea](#mergearea) | [Range](Range.md) | 单元格的合并区域 |
| [MergeCells](#mergecells) | Boolean | 区域内是否存在合并的单元格 |
| [Cells](#cells) | [Range](Range.md) | 区域中的单元格集合 |
| [Rows](#rows) | [Range](Range.md) | 区域中的行集合 |
| [Columns](#columns) | [Range](Range.md) | 区域中的列集合 |
| [EntireRow](#entirerow) | [Range](Range.md) | 区域所在行的整行 |
| [EntireColumn](#entirecolumn) | [Range](Range.md) | 区域所在列的整列 |
| [Row](#row) | Number | 区域中第一行的行号 |
| [RowEnd](#rowend) | Number | 区域中最后一行的行号 |
| [Column](#column) | Number | 区域中第一列的列号 |
| [ColumnEnd](#columnend) | Number | 区域中最后一列的列号 |
| [Borders](#borders) | [Borders](#borders) | 边框集合对象 |

#### 方法列表

| 方法名 | 返回类型 | 简介 |
| --- | --- | --- |
| [BorderAround()](#borderaround) | undefined | 向区域添加边框，并为新边框设置 Border 对象的 Color、LineStyle 和 Weight 属性 |
| [Each()](#each) | undefined | 遍历选区所选单元格 |
| [Item()](#item) | [Range](Range.md) | 表示区域中指定的位置 |
| [Offset()](#offset) | undefined | 对指定区域进行迁移操作 |
| [Replace()](#replace) | undefined | 对单元格内文本执行替换操作 |
| [Delete()](#delete) | undefined | 单元格、行、列的删除 |
| [Insert()](#insert) | undefined | 单元格、行、列的新增 |
| [InsertImage()](#insertimage) | undefined | 插入单元格图片 |
| [Merge()](#merge) | undefined | 合并单元格 |
| [UnMerge()](#unmerge) | undefined | 取消合并单元格 |
| [Address()](#address) | String | 获取表示使用宏语言的区域引用的 String 值 |
| [AddComment()](#addcomment) | undefined | 添加评论 |
| [ClearComments()](#clearcomments) | undefined | 清除区域的评论 |
| [Clear()](#clear) | undefined | 清空指定区域数据和样式 |
| [ClearContents()](#clearcontents) | undefined | 清除区域的内容 |
| [ClearFormats()](#clearformats) | undefined | 清除区域的样式 |
| [ClearHyperlinks()](#clearhyperlinks) | undefined | 清除区域的超链接样式 |
| [Contain()](#contain) | Boolean | 判断区域是否重叠 |
| [Copy()](#copy) | Boolean | 将当前区域对象复制到剪贴板 |
| [Cut()](#cut) | Boolean | 将当前区域对象粘贴到目标区域 |
| [PasteSpecial()](#pastespecial) | undefined | 将剪贴板中的内容粘贴到指定的单元格或范围 |
| [FillLeft()](#fillleft) | undefined | 对指定区域中的单元格执行从右往左填充 |
| [FillRight()](#fillright) | undefined | 对指定区域中的单元格执行从左往右填充 |
| [FillDown()](#filldown) | undefined | 对指定区域中的单元格执行从上往下填充 |
| [FillUp()](#fillup) | undefined | 对指定区域中的单元格执行从下往上填充 |
| [AutoFill()](#autofill) | undefined | 对指定区域中的单元格执行自动填充 |
| [AutoFilter()](#autofilter) | undefined | 对指定区域中的单元格执行自动筛选 |
| [AutoFit()](#autofit) | undefined | 更改区域中的列宽或行高以达到最佳匹配 |
| [Select()](#select) | undefined | 选择区域 |
| [TextToColumns()](#texttocolumns) | undefined | 将包含文本的一列单元格分解为若干列 |

## Count

区域中列/行/单元格的数量，默认为单元格

#### 数据类型

Number - 区域中列/行/单元格的数量

#### 示例

```js
// 默认为单元格的Range对象，Range.Count等价于Range.Cells.Count
let range = Application.Range('A1:B2')
console.log(range.Count) // 4
console.log(range.Cells.Count) // 4
console.log(range.Rows.Count) // 2
console.log(range.Columns.Count) // 2
```

## Text

【只读】读取单元格格式化文本

#### 数据类型

String - 单元格格式化文本

#### 示例

```js
// 区域对象
let range = Application.Range('A1')

// 【只读】读取单元格格式化文本
console.log(range.Text)

// 修改单元格的内容，Text是只读的，修改需要用Value
range.Value = 'NewVaule'
console.log(range.Text) // NewValue
```

## Value/Value2

读写单元格中的值,支持某个区域读取

#### 数据类型

读取单个单元格时：any

读取区域单元格：[][]any

#### 示例

```js
// Value
let range = Application.Range('A1')

range.Value = 'A1'
// 打印A1单元格数据
console.log(range.Value) // A1

range = Range('A1:B2')

// 给A1到B2区域赋值
range.Value = 'WebOffice'

// 打印A1到B2单元格数据
console.log(range.Value) // [["WebOffice","WebOffice"],["WebOffice","WebOffice"]]

// 批量赋值
range.Value = [
  ['A1', 'B1'],
  ['A2', 'B2']
]

// 打印A1到B2单元格数据
console.log(range.Value) // [["A1","B1"],["A2","B2"]]
```

## FormatConditions

FormatConditions 集合对象用于控制 Excel 中的条件格式。条件格式是一种在工作表中格式化单元格的方法，可以根据单元格的值、公式或其他条件自动应用格式，使数据更易于理解和分析

#### 数据类型

[FormatConditions](FormatConditions.md) - FormatConditions 集合对象

#### 示例

```js
// 获取FormatConditions对象
const formatConditions = Range('A:A').FormatConditions
// 获取条件格式对象数量
const count = formatConditions.Count

console.log(count)
```

## Formula

Formula 支持读取和设置单元格内容，但细节有些不同

- 读取情况下：Formula 读取的是公式,而 Text 读取的是值。
- 赋值情况下：Formula 只支持赋值单个单元格,而 Value 支持赋值区域

#### 数据类型

String - 隐式交叉的公式

#### 示例

```js
// 设置A1=1,B1=2
Application.Range('A1').Value = '1'
Application.Range('B1').Value = '2'

// 设置A2单元格的公式: =A1+B1
const range = Application.Range('A2')
range.Formula = '=A1+B1'
// 读取A2单元格的值: 3
console.log(range.Text) // 3
// 读取A2单元格的公式: =A1+B1
console.log(range.Formula) // =A1+B1
```

## FormulaArray

相较于 Formula，支持设置范围数据，其他与 Formula 一致

#### 数据类型

String - 区域的数组公式

#### 示例

```js
let range = Application.Range('A1:B2')

// 设置A1到B2的单元格的值: 100
range.FormulaArray = '=Sum(A1:C3)'

range.Each(function (item) {
  console.log(item.Text) //=Sum(A1:C3)，=Sum(A1:C3)，=Sum(A1:C3)，=Sum(A1:C3)
})
```

## NumberFormat

获取或者设置区域的数字格式。

在获取上，如果指定区域中的所有单元格的数字格式不一致，则此属性返回第一个单元格的数字格式。

在设置上，为了让用户设置更方便，我们列举了 `WebOffice` 上常用的值：

- 常规：`G/通用格式`
- 数值：`0.00_);[红色](0.00)`
- 货币：`￥#,##0.00_);[红色](￥#,##0.00)`
- 会计专用：`_ ￥* #,##0.00_ ;_ ￥* -#,##0.00_ ;_ ￥* "-"??_ ;_ @_`
- 短日期：`yyyy/m/d;@`
- 长日期：`yyyy"年"m"月"d"日";@`
- 时间：`h:mm:ss;@`
- 百分比：`0.00%`
- 分数：`## ?/?`
- 科学技术：`0.00E+00`
- 文本：`@`
- 千位分隔样式：`_ * #,##0.00_ ;_ * -#,##0.00_ ;_ * "-"??_ ;_ @_`

这样在调用 API 设置数字格式的时候，可以复制上面的值进行设置，例如设置数字格式为常规：`Range.NumberFormat = 'G/通用格式'`。

#### 数据类型

String - 区域的数字格式

#### 示例

```js
// 区域对象
let range = Application.Range('A1')

// 获取区域数字格式
console.log('当前区域数字格式：', range.NumberFormat) //当前区域数字格式： G/通用格式

// 设置区域数字格式为文本
range.NumberFormat = '@'
console.log('当前区域数字格式：', range.NumberFormat) //当前区域数字格式： @
```

## Hidden

行或者列的隐藏，该属性只写，无法获取原先的列是否隐藏

#### 数据类型

Boolean - 行或者列的隐藏

#### 示例

```js
// 区域对象
let range = Application.Range('A1')

// 获取整列
let entireColumn = range.EntireColumn

// 隐藏该列
entireColumn.Hidden = true
```

## Interior.Color

区域内部颜色

#### 数据类型

Number - 内部颜色的十六进制 RGB

#### 示例

```js
// 内部属性对象
let interior = Application.Range('A1').Interior

// 获取内部颜色
console.log(interior.Color) // 16777215

// 设置内部颜色为红色
interior.Color = RGB(255, 0, 0)
```

## HorizontalAlignment

设置区域的水平对齐方式，设置的值可以是 [Enum.XlHAlign](Enum.md#xlhalign) 中的值。

#### 数据类型

[Enum.XlHAlign](Enum.md#xlhalign) - 区域的水平对齐方式

#### 示例

```js
// 区域对象
let range = Application.Range('A1')

// 打印对齐方式
console.log(range.HorizontalAlignment) // 1

// 设置对齐方式：居中
range.HorizontalAlignment = Application.Enum.XlHAlign.xlHAlignCenter

// 打印对齐方式
console.log(range.HorizontalAlignment) // -4108，即居中枚举值
```

## VerticalAlignment

设置区域的垂直对齐方式，设置的值可以是 [Enum.XlVAlign](Enum.md#xlvalign) 中的值。

#### 数据类型

[Enum.XlVAlign](Enum.md#xlvalign) - 区域的垂直对齐方式

#### 示例

```js
// 区域对象
let range = Application.Range('A1')

// 设置对齐方式：底部对齐
range.VerticalAlignment = Application.Enum.XlVAlign.xlVAlignBottom

// 获取对齐方式
console.log(range.VerticalAlignment) //-4107，即向下垂直对齐方式枚举值
```

## IndentLevel

设置单元格缩进

#### 数据类型

Number - 单元格缩进

#### 示例

```js
// 区域对象
let range = Application.Range('A1')

// 写入值到单元格中
console.log(range.IndentLevel) // 0
range.IndentLevel = 10
console.log(range.IndentLevel) // 10
```

## WrapText

获取或者设置区域自动换行

#### 数据类型

Boolean - 区域是否自动换行

#### 示例

```js
// 区域对象
let range = Application.Range('A1')
// 获取区域自动换行
console.log('区域是否自动换行：', range.WrapText) // true
range.WrapText = false
console.log('区域是否自动换行：', range.WrapText) // false
```

## MergeArea

单元格的合并区域

#### 数据类型

Range - 该区域内合并单元格的 Range 对象

#### 示例

```js
// 区域对象
let range = Application.Range('A1:D2')
// 合并单元格
range.Merge()
// 将该合并区域值设置为merge
range.MergeArea.Value = 'merge'
```

## MergeCells

区域内是否存在合并的单元格

#### 数据类型

Boolean - 是否存在合并的单元格

#### 示例

```js
// 区域对象
let range = Application.Range('A1:D2')
// 合并单元格
range.Merge()
// 区域内是否存在合并的单元格
console.log(range.MergeCells) // true
```

## Cells

区域中的所有单元格集合，返回一个 [Range](Range.md) 对象（可使用 Range 相关的属性、方法）

#### 数据类型

[Range](Range.md) - 区域中的所有单元格集合

#### 示例

```js
// 区域对象
let range = Application.Range('A1:B2')

// 单元格对象
let cells = range.Cells

// 取第一个单元格
let cell = cells.Item(1)
console.log(cell.Address()) // $A$1
```

## Rows

区域中的行，返回一个 Range 对象（可使用 Range 相关的属性、方法）

#### 数据类型

[Range](Range.md) - 区域中的所有行集合

#### 示例

```js
// A1:B2最一共有2行，因此返回2
console.log(Application.Range('A1:B2').Rows.Count) // 2

// A1:D4最一共有4行，因此返回4
console.log(Application.Range('A1:D4').Rows.Count) // 4
```

## Columns

区域中的所有列，返回一个 [Range](Range.md) 对象（可使用 Range 相关的属性、方法）

#### 数据类型

[Range](Range.md) - 区域中的所有列集合

#### 示例

```js
// A1:B2最一共有2列，因此返回2
console.log(Application.Range('A1:B2').Columns.Count) // 2

// A1:D4最一共有4列，因此返回4
console.log(Application.Range('A1:D4').Columns.Count) // 4
```

## EntireRow

包含指定区域的整行

#### 数据类型

[Range](Range.md) - 包含指定区域的整行

#### 示例

```js
// 一行默认长度为16384
console.log(Application.Range('A1').EntireRow.Count) // 16384

// 两行
console.log(Application.Range('A1:B2').EntireRow.Count) // 32768
```

## EntireColumn

包含指定区域的整列

#### 数据类型

[Range](Range.md) - 包含指定区域的整列

#### 示例

```js
// 一列默认长度为1048576
console.log(Application.Range('A1').EntireColumn.Count) // 1048576

// 两列
console.log(Application.Range('A1:B2').EntireColumn.Count) // 2097152
```

## Row

区域中第一行的行号

#### 数据类型

Number - 区域中第一行的行号

#### 示例

```js
let range = Application.Range('A1:D4')
// A1:D4第一行是1行，因此返回1
console.log(range.Row) // 1
```

## RowEnd

区域中最后一行的行号

#### 数据类型

Number - 区域中最后一行的行号

#### 示例

```js
let range = Application.Range('A1:D4')
// A1:D4最后一行是4行，因此返回4
console.log(range.RowEnd) // 4
```

## Column

区域中最左列的列号

#### 数据类型

Number - 区域中最左列的列号

#### 示例

```js
let range = Application.Range('A1:D4')
// A1:D4最左列是A列，因此返回1
console.log(range.Column) // 1
```

## ColumnEnd

区域中最右列的列号

#### 数据类型

Number - 区域中最右列的列号

#### 示例

```js
let range = Application.Range('A1:D4')
// A1:D4最右列是D列，因此返回4
console.log(range.ColumnEnd) // 4
```

## Borders

边框对象

#### 数据类型

[Border](Border.md)集合对象 - 边框对象集合

#### 示例

```js
// 边框对象
let borders = Application.Range('A1').Borders

// 将A1的左上角到右下角的边框颜色设置为红色
borders.Item(Application.Enum.XlBordersIndex.xlDiagonalDown).Color = '#FF0000'
```

### Borders.Item()

单个边框对象，代表单元格区域或样式的边框之一

#### 参数

| 属性 | 数据类型 | 默认值 | 必填 | 说明 |
| --- | --- | --- | --- | --- |
| Index | Enum |  | 是 | 指定要检索的边框，参考 [Enum.XlBordersIndex](Enum.md#xlbordersindex) |

#### 返回类型

[Border](#border) - 单个边框对象，代表单元格区域或样式的边框之一

#### 示例

```js
// 边框对象
let borders = Application.Range('A1').Borders

// 将A1的左上角到右下角的边框颜色设置为黄色
borders.Item(Application.Enum.XlBordersIndex.xlDiagonalDown).Color = '#FFFF00'
```

## BorderAround()

向区域添加边框，并为新边框设置 Border 对象的 Color、LineStyle 和 Weight 属性

#### 参数

| 属性 | 数据类型 | 默认值 | 必填 | 说明 |
| --- | --- | --- | --- | --- |
| LineStyle | [LineStyle](Enum.md#xllinestyle) | null | 否 | XlLineStyle 的常量之一，指定边框的线条样式 |
| BorderWeight | [BorderWeight](Enum.md#xlborderweight) | null | 否 | 边框粗细 |
| Color | string | null | 否 | 边框颜色，以 RGB 值表示，例如`#ff0000` |

#### 示例

```js
// 为B2单元格设置红色虚线
Range('B2').BorderAround(xlDash, xlHairline, '#ff0000')
// 为C3单元格设置绿色点划相间线
Range('C3').BorderAround(xlDashDot, xlHairline, '#00ff00')
```

## Each()

遍历选区所选单元格，建议使用此函数时不要涉及插入或删除行列，否则可能会导致不符合预期的结果

#### 参数

| 属性 | 数据类型 | 默认值 | 必填 | 说明 |
| --- | --- | --- | --- | --- |
| callback | Function | null | 是 | 类似 JS 数组的 forEach |

#### 示例

```js
// 区域对象
let range = Application.Range('A1:D2')

// 编辑选区单元格
range.Each(function (item) {
  //打印单元格的地址
  console.log(item.Address()) // $A$1...$D$2
})
```

## Item()

返回一个 Range 对象，表示区域中指定的位置

#### 参数

| 属性 | 数据类型 | 默认值 | 必填 | 说明 |
| --- | --- | --- | --- | --- |
| RowIndex | Number |  | 是 | 如果提供了第二个参数，则返回的单元格的相对行号。如果未提供第二个参数，则为要返回的子范围的索引 |
| ColumnIndex | Number |  | 否 | 要返回的单元格的相对列号 |

#### 返回类型

Range - 对应的单元格对象

#### 示例

```js
// 区域对象
const range = Application.Range('A1:D2')
// 区域子项：B2
console.log(range.Item(2, 2).Address()) //$B$2
```

## Offset()

对指定区域进行偏移，返回偏移后的 Range

#### 参数

| 属性 | 数据类型 | 默认值 | 必填 | 说明 |
| --- | --- | --- | --- | --- |
| RowOffset | Number | 0 | 否 | 区域偏移的行数：可以是正值、负值或零。正值表示向下偏移，负值表示向上偏移 |
| ColumnOffset | Number | 0 | 否 | 区域偏移的列数：可以是正值、负值或零。正值表示向右偏移，负值表示向左偏移 |

#### 示例

```js
// 区域对象
let range = Application.Range('A1:D2')
// 区域子项：B2
console.log(range.Address()) //$A$1:$D$2
// 对指定区域进行偏移
let newRange = range.Offset(2, 2)
// 区域子项：B2
console.log(newRange.Address()) //$C$3:$F$4
```

## Replace()

对单元格内文本执行替换操作

#### 参数

| 属性 | 数据类型 | 默认值 | 必填 | 说明 |
| --- | --- | --- | --- | --- |
| What | String | 无 | 是 | 希望搜索的字符串。 |
| Replacement | String | 无 | 是 | 替换字符串。 |
| LookAt | Enum | xlWhole | 否 | 可以是下列 XlLookAt 常量之一：xlWhole 或 xlPart。详情看示例 |
| SearchOrder | Enum | xlByRows | 否 | 可以是以下 XlSearchOrder 常量之一：xlByRows 或 xlByColumns。 详情看示例 |
| MatchCase | Boolean | false | 否 | 如果为 True，则搜索区分大小写。 |
| MatchByte | Boolean | false | 否 | 如果是 true，区分全半角符号 |

#### 示例

```js
// 区域对象
let range = Application.Range('A1:D2')
// 把所有小山全部替换成大山
range.Replace('小山', '大山', xlPart)
// 单元格内容为小山的单元格替换成大山
range.Replace('小山', '大山', xlWhole)
// 把JinXiaoShan替换成金小山，区分大小写
range.Replace('JinXiaoShan', '金小山', xlPart, xlByRows, true)
// 把JinXiaoShan替换成金小山，不区分大小写
range.Replace('JinXiaoShan', '金小山', xlPart, xlByRows, false)
// 把JinXiaoShan替换成金小山，区分全角半角
range.Replace('JinXiaoShan', '金小山', xlPart, xlByRows, true, true)
// 把JinXiaoShan替换成金小山，不区分全角半角
range.Replace('JinXiaoShan', '金小山', xlPart, xlByRows, true, false)
```

## Delete()

单元格、行、列的删除.删除单元格时,右侧单元格左移

#### 示例：删除行

```js
// 区域对象
let range = Application.Range('B5:D10')
// 选择区域的所有行并删除，例子删除5-10行
range.EntireRow.Delete()
```

#### 示例：删除列

```js
let range = Application.Range('B5:D10')
// 选择区域的所有行并删除，例子删除2-4列
range.EntireColumn.Delete()
```

## Insert()

单元格、行、列的新增

#### 示例：新增行

```js
let range = Application.Range('B5:D10')
// 选择区域的所有行并在最上行新增一行，例子在第5行上方插入一行，原5行下移为6行
range.EntireRow.Insert()
```

#### 示例：新增列

```js
let range = Application.Range('B5:D10')
// 选择区域的所有列并在最左列新增一列，例子在第2列左侧插入一列，原2列右移为3列
range.EntireColumn.Insert()
```

## InsertImage()

插入单元格图片

#### 参数

| 属性 | 数据类型 | 默认值 | 必填 | 说明 |
| --- | --- | --- | --- | --- |
| dataURL | string | undefined | 是 | base64 字符串形式的图片 |

#### 示例

```js
// 获取E1单元格
const range = Range('E1')
// 向目标单元格插入图片
range.InsertImage(
  'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAACAAAAAgCAMAAABEpIrGAAACH1BMVEUAAAA9kP8mpv9Fv/8QT94TXeQAfv8trP0xr/0VUdoAcv8LgPs1sf4AgP4QUuURTuEVl/ksq/wAg/84tf4NUt80sf0bnPsPUuAAYfsAgP4UTtYAY/8GXPQRkv1Jxv8AZP86tv9Bvv9Gv/9Atv8KiPUOVesAZ+wbS8wUl/kjo/sAYv8PVekjovxBu/8AY/8trPxCu/8AW/AAcvAxsP0npvsYTdElo/oQT9o+uP8AYv8KVOIHdugAYfwAfPwbfe49uf4XmfkcdOopo/oAYv86tf4AYf8TT9cZaeYVTdUJVOUcnvsAYv8Agv9Iwv8AYv9Jwf8bSs0gofoAY/8Ag/9Gv/8en/oAYf9Fv/8AWuQZS80AdO5BvP8AYv9DvP8Umfo9uv8LVOUAgf8amfk/uf8AY/8Ag/8Ahf8TUdgVT9UNWOknov8trPwxr/wmpvspqfs1sv0io/s5tf0AgP8bnfoMV+4AYPoAYv8AYf4AYfwAe/sen/oAd/gAXvYAff0LW/MLWfEOUN0PT9oSTtZAu/49uP0qqfweoPoXmvkAdPUAXPIAcfEKUuIWTNIXS88KXPYAWusOVeoPVOkPUuYQUeQQT+IMUeAUTdMZSs09uP4Aa+kAWOgAaeQIVOQIU+QAV+MRTd8STd0RTtgSl/kAW+8NWO8Abu0AbuwBZNsAg/8AYPwAc/MIVecAZuEAVeAAWt8AZN0JgvAIfOsJV+sTVuIAU9qil9AQAAAAa3RSTlMAAwb++A3+bkkkGxL8+Pf36tnTubCYjol4dmVjRjEvLysZFA7+/Pv6+fj39/b19fPw8O/u7Ovq6Ofn5+Xj4uDa2trZ1dPS0s/Ny7+9vLi4trapoqGgoJ2XkI+KgoJycGRhX1BNS0pHQj06IcCB3jkAAAIBSURBVDjLdZCHctpAFEWfCMU1ce8tTu+99957770nJFGwkE1sgyMTAginYbkAIQSD494+0G+ltTAacWZnNLt7dO/Og3Tqzl2og8wsvroAudEI+iy5s+6LzMZ7Bp1r5sFOQegRenAJwoEnoCXr6Mc0ihvSrhuK32vJK21Mva00760OWyz0KZYNbzJQmKXUF1o7rZ24rGngweETIGOwbH2nw/ZDUaM6grJVHzSs3R+NoqDy+lKLuwWX2618r03/QRalhNzbewo+qZyuZX4TVMFQubm3f3xy5VeZg49xsL8IC+mUH+1zeLyJ3pHkRCwWK7hvIGc/CYrwvIhtbXZ4hgYwJLm0LFf56QdhDRAq8wdZMypD3kT/SBJo6nfCMmWTc2WQlUO8GDIn2GSAUnuWGERJpIS4LR4HlWdFNITua9pl5s3h5vJ81tzscAAhuyTc3kZQ51CxqasvMmDGENyZyo1BX2iqLRAI0L7qvU2fv6Ex5mVbgana4XIGfb7QsI0KNcc4O4/Gv67I6Jjn5cmOvy4nGuFwaFipqFjN2YlBQkZn/N0dxAhiSfjUKzqHyxwq/0lIX0T0d6OAGb7dD0Gl/gzH8UqN6JcNp7HcBPN5esTO8U2oSCIpcZVkgwamepedxxpJROP4C9DBdHc9vlWSxG1VDOiTc53jpRW3TJCZ+vMXNeWzDz4DoNZyqecAAAAASUVORK5CYII='
)
```

## Merge()

合并单元格

#### 参数

| 属性 | 数据类型 | 默认值 | 必填 | 说明 |
| --- | --- | --- | --- | --- |
| Across | boolean | false | 是 | 如果设置为 true，则将指定区域中每一行的单元格合并为一个单独的合并单元格 |

#### 示例

```js
const range = Application.Range('A1:D2')
// 合并单元格
range.Merge()
```

## UnMerge()

取消合并单元格

#### 参数

| 属性 | 数据类型 | 默认值 | 必填 | 说明 |
| --- | --- | --- | --- | --- |
| CancelCenter | Boolean | false | 否 | 是否合并居中 |

#### 示例

```js
const range = Application.Range('A1:D2')

// 取消合并单元格
range.UnMerge()
```

## Address()

获取表示使用宏语言的区域引用的 String 值

#### 参数

| 属性 | 数据类型 | 默认值 | 必填 | 说明 |
| --- | --- | --- | --- | --- |
| RowAbsolute | Boolean |  | 否 | 若为 true，以绝对引用的形式返回引用的行部分。默认值为 true |
| ColumnAbsolute | Boolean |  | 否 | 若为 true，以绝对引用的形式返回引用的列部分。默认值为 true |
| ReferenceStyle | Enum |  | 否 | 引用样式。默认值为 xlA1，更多可看 [Enum.XlReferenceStyle](Enum.md#xlreferencestyle) |
| External | Boolean |  | 否 | 若为 true，返回外部引用。若为 false，返回本地引用。默认值为 false |
| RelativeTo | Range |  | 否 | 如果 RowAbsolute 和 ColumnAbsolute 为 false，且 ReferenceStyle 是 xlR1C1，则必须为相对引用包含一个起点。此参数是一个定义起点的 Range 对象 |

#### 返回类型

String - 表示使用宏语言的区域引用的值

#### 示例

```js
// 区域对象
let range = Application.Range('A1')

// 获取宏语言的区域引用的 String 值
console.log('address1：', range.Address())

console.log('address2：', range.Address(false, false))

console.log(
  'address3：',
  range.Address(true, true, Application.Enum.XlReferenceStyle.xlR1C1)
)
```

## AddComment()

添加评论

#### 参数

| 属性 | 数据类型 | 默认值 | 必填 | 说明 |
| --- | --- | --- | --- | --- |
| Text | String |  | 否 | 评论文本 |

#### 示例

```js
let range = Application.Range('A1')
// 给 A1 区域添加评论
range.AddComment('WebOffice')
```

## ClearComments()

清除区域的评论

#### 示例

```js
let range = Application.Range('A1')
// 清除区域的评论
range.ClearComments()
```

## Clear()

清空指定区域数据和样式

#### 示例

```js
const range = Range('A1:B2')
range.Value = [
  ['A1', 'B1'],
  ['A2', 'B2']
]
range.Clear()
console.log(range.Value) //[["",""],["",""]]
```

## ClearContents()

清除区域的内容

#### 示例

```js
let range = Application.Range('A1')
// 清除区域的内容
range.ClearContents()
```

## ClearFormats()

清除区域的样式

#### 示例

```js
let range = Application.Range('A1')
// 清除区域的样式
range.ClearFormats()
```

## ClearHyperlinks()

清除区域的超链接样式

#### 示例

```js
let range = Application.Range('A1')
// 清除区域的超链接样式
range.ClearHyperlinks()
```

## Contain()

判断区域是否重叠

#### 参数

| 属性 | 数据类型 | 默认值 | 必填 | 说明 |
| --- | --- | --- | --- | --- |
| Range | Range |  | 是 | Range 对象，另一块区域 |

#### 返回类型

Boolean - 区域是否重叠

#### 示例

```js
// 区域对象
let range = Application.Range('A1:D2')

// 第二块区域对象
let newRange = Application.Range('A1:B4')

// 判断是否重叠
console.log(range.Contain(newRange))
```

## Copy()

将当前区域对象复制到剪贴板

#### 返回类型

Boolean - 是否复制成功

#### 示例

```js
// 复制A1
Range('A1').Copy()
// 粘贴到B1
Range('B1').PasteSpecial()
```

## Cut()

将当前区域对象粘贴到目标区域

#### 参数

| 属性 | 数据类型 | 默认值 | 必填 | 说明 |
| --- | --- | --- | --- | --- |
| Destination | [Range](Range.md) | null | 是 | 目标区域对象 |

#### 返回类型

Boolean - 是否复制成功

#### 示例

```js
// 将A1区域的数据粘贴到A2区域
Range('A1').Cut(Range('A2'))
```

## PasteSpecial()

用于将剪贴板中的内容粘贴到指定的单元格或范围。它可以以多种方式粘贴数据，包括值、格式、公式等。可以通过指定参数来控制粘贴的方式，如粘贴值、粘贴格式、粘贴公式等

注意

在使用 PasteSpecial()粘贴数据之前，请确保已将源区域的值复制到剪贴板，如果剪贴板内没有任何区域对象，则会抛出剪贴板函数调用异常。

您可先通过[Copy()](#copy)函数将区域对象复制到剪贴板，然后再使用 PasteSpecial()函数进行粘贴，并可以选择性地应用不同的粘贴选项。

#### 参数

| 属性 | 数据类型 | 默认值 | 必填 | 说明 |
| --- | --- | --- | --- | --- |
| Paste | [XlPasteType](Enum.md#xlpastetype) | xlPasteAll | 否 | 粘贴类型，例如 xlPasteAll 或 xlPasteValues |
| Operation | [XlPasteSpecialOperation](Enum.md#xlpastespecialoperation) | xlPasteSpecialOperationNone | 否 | 粘贴操作，例如 xlPasteSpecialOperationAdd |
| SkipBlanks | Boolean | false | 否 | 如果为 true，则不将剪贴板上区域中的空白单元格粘贴到目标区域中 |
| Transpose | Boolean | false | 否 | 如果为 true ，则表示在粘贴区域时转置行和列 |

#### 返回类型

Undefined

#### 示例一

粘贴值：使用此方法以值形式将剪贴板中的内容粘贴到指定单元格或范围。例如，将 B1 的值粘贴到 A1

```js
Range('B1').Copy()
Range('A1').PasteSpecial(xlPasteValues)
```

#### 示例二

粘贴格式：将 A1:B2 区域的格式粘贴至 C1:D2

```js
Range('A1:B2').Copy()
Range('C1:D2').PasteSpecial(xlPasteFormats)
```

#### 示例三

粘贴值和格式：将 A1:B2 区域的值和格式粘贴至 A3:B4

```js
Range('A1:B2').Copy()
Range('A3:B4').PasteSpecial(xlPasteAll)
```

#### 示例四

转置粘贴：将 A1:B2 区域的值转置粘贴至 C1:D2（转置粘贴后，行和列会互换）

```js
Range('A1:B2').Copy()
Range('C1:D2').PasteSpecial(
  xlPasteValues,
  xlPasteSpecialOperationNone,
  false,
  true
)
```

## FillLeft()

对指定区域中的单元格执行从右往左填充,填充值为最右列的值

#### 示例

```js
const range = Range('A1:B2')
range.Value = [
  ['A1', 'B1'],
  ['A2', 'B2']
]
range.FillLeft()
console.log(range.Value) // [["B1","B1"],["B2","B2"]]
```

## FillRight()

对指定区域中的单元格执行从左往右填充,填充值为最左列的值

#### 示例

```js
const range = Range('A1:B2')
range.Value = [
  ['A1', 'B1'],
  ['A2', 'B2']
]
range.FillRight()
console.log(range.Value) // [["A1","A1"],["A2","A2"]]
```

## FillDown()

对指定区域中的单元格执行从上往下填充,填充值为最上列的值

#### 示例

```js
const range = Range('A1:B2')
range.Value = [
  ['A1', 'B1'],
  ['A2', 'B2']
]
range.FillDown()
console.log(range.Value) // [["A1","B1"],["A1","B1"]]
```

## FillUp()

对指定区域中的单元格执行从下往上填充,填充值为最下列的值

#### 示例

```js
const range = Range('A1:B2')
range.Value = [
  ['A1', 'B1'],
  ['A2', 'B2']
]
range.FillUp()
console.log(range.Value) //[["A2","B2"],["A2","B2"]]
```

## AutoFill()

对指定区域中的单元格执行自动填充

#### 参数

| 属性 | 数据类型 | 默认值 | 必填 | 说明 |
| --- | --- | --- | --- | --- |
| Destination | Range |  | 否 | 目标区域。目标区域必须包含源区域 |
| Type | Enum | Enum.XlAutoFillType.xlFillDefault | 否 | 填充类型，详细可见 [Enum.XlAutoFillType](Enum.md#xlautofilltype) |

#### 示例

```js
// 区域对象
let range = Application.Range('A1')
range.Value = 1
// 要填充的单元格
let fillRange = Application.Range('A1:A20')
// 对指定区域中的单元格执行自动填充
range.AutoFill(fillRange)
fillRange.Each(item => {
  console.log(item.Text()) // 自动填充为：1,2,3,4...20
})
```

## AutoFilter()

对指定区域的单元格执行自动筛选

#### 参数

| 属性 | 数据类型 | 默认值 | 必填 | 说明 |
| --- | --- | --- | --- | --- |
| Field | Variant | null | 否 | 指定想要基于筛选的字段的整数偏移量。从列表的左侧算起，最左侧的字段是 1 |
| Criteria1 | Variant | null | 否 | 指定判断条件。使用“=”查找空字段，或者使用“<>”查找非空字段。如果忽略该参数，那么判断是全部。如果参数 Operator 是 xlTop10Items，那么参数 Criterial1 指定项目的数量 |
| Operator | [XlAutoFilterOperator](Enum.md#xlautofilteroperator) | null | 否 | 指定筛选的类型，通过枚举值[XlAutoFilterOperator](Enum.md#xlautofilteroperator)来指定 |
| Criteria2 | Variant | null | 否 | 第二个判断条件。与 Criteria1 和 Operator 一起组合成复合筛选条件。 也用作日期字段的单一条件（按日、月或年筛选）。 后跟一个数组，该数组用于描述筛选 Array(Level, Date)。 其中，Level 为 0-2（年、月、日），Date 为筛选期内的一个有效日期 |
| SubField | Variant | null | 否 | 对其应用条件的数据类型中的字段（例如，来自地理位置的“人口”字段或来自股票的“交易量”字段）。省略此值目标是“（显示值）” |
| VisibleDropDown | Variant | true | 否 | 如果为 `true`，则显示已筛选字段的 `AutoFilter` 下拉箭头。 如果为 `false`，则隐藏已筛选字段的 `AutoFilter` 下拉箭头。 默认情况下为 `true` |

#### 示例一

将 B 列数值排在后面的 11 行筛选显示

```js
// 未定义筛选区域时，默认使用UsedRange作为筛选区域
const filterRange = ActiveSheet.UsedRange
// 定义筛选区域B列
let filterField = 2
// 获取筛选的数据区域
const filterColumnRange = filterRange.Columns(filterField)
// 计算数据区域的行数，需要减去表头
const filterRangeCount = filterColumnRange.Rows.Count - 1
// 获取后11名数据
const bottom11Value = 11
// 数据区域中可能没有11行数据
const bottomNum =
  filterRangeCount < bottom11Value ? filterRangeCount : bottom11Value
// 获取筛选的数据里第11小的值
let bottom11 = WorksheetFunction.Small(filterColumnRange, bottomNum)
// 应用筛选
filterRange.AutoFilter(filterField, '<=' + bottom11)
```

#### 示例二

将第 1 列为"1 班"，第 5 列的值小于 60 的单元格填充为红色，第 5 列的值大于 90 的单元格填充为绿色

```js
// 选择筛选区域为已使用的区域
const filterRange = ActiveSheet.UsedRange
// 设置筛选第1列
const filterColumn1 = 1
// 设置筛选值为"1班"
let filterValue = '1班'
// 应用筛选
filterRange.AutoFilter(filterColumn1, filterValue)
// 选择要筛选的列号为第5列
const filterColumn2 = 5
// 选择筛选的值为小于"60"
filterValue = '<60'
// 应用筛选
filterRange.AutoFilter(filterColumn2, filterValue)
// 获取除标题行外的数据区域
const dataRange = ActiveSheet.UsedRange.Offset(1).Resize(
  ActiveSheet.UsedRange.Rows.Count - 1
)
// 定义列的区域
const columnRange = dataRange.Columns(filterColumn2)
// 获取筛选后显示的数据
const visibleRange = columnRange.SpecialCells(xlCellTypeVisible)
// 定义颜色红色
const redColor = RGB(255, 0, 0)
// 设置颜色为红色
visibleRange.Interior.Color = redColor
// 清除第5列的筛选条件
filterRange.AutoFilter(filterColumn2)
// 选择筛选的值为大于"90"
const filterValue2 = '>90'
// 应用筛选
filterRange.AutoFilter(filterColumn2, filterValue2)
// 获取筛选后显示的数据
const visibleRange2 = columnRange.SpecialCells(xlCellTypeVisible)
// 定义颜色绿色
const greenColor = RGB(0, 255, 0)
// 设置颜色为绿色
visibleRange2.Interior.Color = greenColor
// 清除第1列的筛选条件
filterRange.AutoFilter(filterColumn1)
// 清除第5列筛选条件
filterRange.AutoFilter(filterColumn2)
```

#### 示例三

筛选出 A 列是 2023 年 4 月 10 日和 2022 年 12 月并且 C 列值高于 10000 的数据/信息/情况/商品/行

```js
// 未定义筛选区域时，默认使用UsedRange作为筛选区域
const filterRange = ActiveSheet.UsedRange
// 定义筛选区域A列
let filterField = 1
// 定义日期筛选时间
let criteria = new Array()
// 定义筛选年、月、日的枚举值
const FilterDateTimeEnum = { YEAR: 0, MONTH: 1, DAY: 2 }
// 增加筛选日期为2023年
criteria.push(FilterDateTimeEnum.DAY, '2023/4/10')
// 增加筛选日期为2022年12月
criteria.push(FilterDateTimeEnum.MONTH, '2022/12/1')
// 应用筛选
filterRange.AutoFilter(filterField, undefined, xlFilterValues, criteria)
// 定义筛选区域为C列
filterField = 3
// 设置筛选条件为大于10000
let criteria2 = '>10000'
// 应用筛选
filterRange.AutoFilter(filterField, criteria2)
```

#### 示例四

筛选出 A 列日期是 2023 年、2022 年 3 月、2022 年 4 月、2022 年 5 月 1 日、2022 年 5 月 2 日的数据

```js
// 未定义筛选区域时，默认使用UsedRange作为筛选区域
const filterRange = ActiveSheet.UsedRange
// 定义筛选区域A列
let filterField = 1
// 定义日期筛选时间
let criteria = new Array()
// 定义筛选年、月、日的枚举值
const FilterDateTimeEnum = { YEAR: 0, MONTH: 1, DAY: 2 }
// 增加筛选日期为2023年
criteria.push(FilterDateTimeEnum.YEAR, '2023/1/1')
// 增加筛选日期为2022年3月
criteria.push(FilterDateTimeEnum.MONTH, '2022/3/1')
// 增加筛选日期为2022年4月
criteria.push(FilterDateTimeEnum.MONTH, '2022/4/1')
// 增加筛选日期为2022年5月1日
criteria.push(FilterDateTimeEnum.DAY, '2022/5/1')
// 增加筛选日期为2022年5月2日
criteria.push(FilterDateTimeEnum.DAY, '2022/5/2')
// 应用筛选
filterRange.AutoFilter(filterField, undefined, xlFilterValues, criteria)
```

#### 示例五

帮我将 D 列是本季度同时 C 列是去年到现在截止时间为一年，A 列是'商品 A'显示出来

```js
// 未定义筛选区域时，默认使用UsedRange作为筛选区域
const filterRange = ActiveSheet.UsedRange
// 定义筛选区域D列
let filterField = 4
// 设置筛选 xlFilterThisQuarter， 可选值xlFilterLastQuarter、xlFilterNextQuarter、xlFilterThisQuarter
let criteria = xlFilterThisQuarter
// 应用筛选
filterRange.AutoFilter(filterField, criteria, xlFilterDynamic)
// 定义筛选字段为C列
filterField = 3
// 设置筛选时间为 xlFilterYearToDate 过去到今天为止一年的时间
criteria = xlFilterYearToDate
// 应用筛选
filterRange.AutoFilter(filterField, criteria, xlFilterDynamic)
// 定义筛选区域为A列
filterField = 1
// 设置筛选条件'商品A'
criteria = '商品A'
// 应用筛选
filterRange.AutoFilter(filterField, criteria)
```

#### 示例六

筛选第二列包含'张三'或者开头等于'李四'的数据

```js
// 未定义筛选区域时，默认使用UsedRange作为筛选区域
const filterRange = ActiveSheet.UsedRange
// 定义筛选区域第2列
let filterField = 2
// 设置筛选包含内容"张三"
let criteria = '=*张三*'
// 设置筛选开头是"李四"
let criteria2 = '=李四*'
// 应用筛选，指定包含关系
filterRange.AutoFilter(filterField, criteria, xlOr, criteria2)
```

#### 示例七

将 A 列数值排在前 5 的行筛选显示

```js
// 未定义筛选区域时，默认使用UsedRange作为筛选区域
const filterRange = ActiveSheet.UsedRange
// 定义筛选区域A列
let filterField = 1
// 获取筛选的数据区域
const filterColumnRange = filterRange.Columns(filterField)
// 计算数据区域的行数，需要减去表头
const filterRangeCount = filterColumnRange.Rows.Count - 1
// 获取前5名数据
const top5Value = 5
// 数据区域中可能没有5行数据
const topNum = filterRangeCount < top5Value ? filterRangeCount : top5Value
// 获取筛选的数据里第5大的值
let top5 = WorksheetFunction.Large(filterColumnRange, topNum)
// 应用筛选
filterRange.AutoFilter(filterField, '>=' + top5)
```

#### 示例八

将 C 列为"张三"的 F 列设置为楷体字体

```js
// 选择筛选区域为已使用的区域
const filterRange = ActiveSheet.UsedRange
// 选择要筛选的列号为3
const filterColumn = 3
// 选择筛选的值为"张三"
const filterValue = ['张三']
// 应用筛选
filterRange.AutoFilter(filterColumn, filterValue)

// 获取筛选结果
const filteredRange =
  ActiveSheet.AutoFilter.Range.SpecialCells(xlCellTypeVisible)

// 获取筛选结果的F列
const filteredRangeFColumn =
  ActiveSheet.AutoFilter.Range.Columns('F').SpecialCells(xlCellTypeVisible)

// 将筛选的结果的F列设置为楷体字体
filteredRangeFColumn.Font.Name = '楷体'
```

#### 示例九

筛选仅显示以下符合条件的内容：第 2 列内容为'李四’并且第 4 列是上周

```js
// 未定义筛选区域时，默认使用UsedRange作为筛选区域
const filterRange = ActiveSheet.UsedRange
// 定义筛选区域第2列
let filterField = 2
// 设置筛选内容为"李四"
let criteria = '李四'
// 应用筛选
filterRange.AutoFilter(filterField, criteria)
// 定义筛选区域为第4列
filterField = 4
// 设置筛选条件为上周，可选值 xlFilterLastWeek，xlFilterThisWeek，xlFilterNextWeek
criteria = xlFilterLastWeek
// 应用筛选
filterRange.AutoFilter(filterField, criteria, xlFilterDynamic)
```

## AutoFit()

更改区域中的列宽或行高以达到最佳匹配

被操作的对象必须是[Columns](#column)或者[Rows](#rows)，否则，该方法将不会有效果。

一个列宽单位等于“常规”样式中一个字符的宽度

#### 示例

```js
// A列设置自动列宽
const range = Application.Range('A1')
range.Columns.AutoFit()

// 第一行设置自动行高
const range = Application.Range('A1')
range.Rows.AutoFit()
```

## Select()

选择区域，通过该调用可以修改 Application.Selection 的选区访问

#### 示例

```js
// 区域对象
let range = Application.Range('A1')
range.Select()
let selection = Application.Selection
console.log(selection.Address())
```

## TextToColumns()

将单元格中的文本按指定的分隔符分成多个列

#### 参数

| 属性 | 数据类型 | 默认值 | 必填 | 说明 |
| --- | --- | --- | --- | --- |
| Destination | [Range](Range.md) | null | 否 | 指定要将结果插入到的单元格或单元格区域。如果省略此参数，则结果将覆盖源单元格 |
| DataType | [XlTextParsingType](Enum.md#xltextparsingtype) | xlDelimited | 否 | 指定将被拆分到多列中的文本的格式 |
| TextQualifier | [XlTextQualifier](Enum.md#xltextqualifier) | xlTextQualifierDoubleQuote | 否 | 指定文本分隔符 |
| ConsecutiveDelimiter | Boolean | false | 否 | 指定是否将连续分隔符视为一个分隔符 |
| Tab | Boolean | false | 否 | 指定是否使用制表符作为分隔符 |
| Semicolon | Boolean | false | 否 | 指定是否使用分号作为分隔符 |
| Comma | Boolean | false | 否 | 指定是否使用逗号作为分隔符 |
| Space | Boolean | false | 否 | 指定是否使用空格作为分隔符 |
| Other | Boolean | false | 否 | 指定是否使用 `OtherChar` 字符的内容作为分隔符 |
| OtherChar | String | `''` | 否 | 指定自定义分隔符，默认值为空字符串 |
| FieldInfo | Array | null | 否 | 包含各个数据列解析信息的数组 |
| DecimalSeparator | String | 系统设置 | 否 | 识别数字时，Excel 使用的小数分隔符 |
| ThousandsSeparator | String | 系统设置 | 否 | 识别数字时，Excel 使用的千位分隔符 |
| TrailingMinusNumbers | String | null | 否 | 以减号字符开始的数字 |

提示

关于`FieldInfo`字段：此数组包含各个数据列解析信息，具体取决于`DataType`的值。

分隔数据时，此参数是一个双元素数组，每个双元素数组指定特定列的转换选项。 第一个元素是列号 (从 1 开始的) ，第二个元素是指定应该如何解析列，取值是 [XlColumnDataType](Enum.md#xlcolumndatatype) 枚举。如果输入数据中特定列不存在给定列说明符，则使用 “常规 ”设置分析该列。如果源数据具有固定宽度的列，则每个双元素数组的第一个元素将列的起始字符位置指定为整数。

#### 返回类型

undefined

#### 示例一

以 tab 为分隔符分隔 A1:A3 里面的内容。

以分号为分隔符分隔 B 列。

以逗号为分隔符分隔 C1。

以空格为分隔符分离 D1 里面的内容。

```js
// 1. 以tab为分隔符拆分A1:A3，Tab表示用tab来分隔
Range('A1:A3').TextToColumns(
  undefined,
  xlDelimited,
  xlTextQualifierNone,
  undefined,
  true,
  undefined,
  undefined,
  undefined,
  undefined,
  undefined,
  undefined
)
// 2. 以分号为分隔符拆分B列，Semicolon表示用分号来分隔
Range('B:B').TextToColumns(
  undefined,
  xlDelimited,
  xlTextQualifierNone,
  undefined,
  undefined,
  true,
  undefined,
  undefined,
  undefined,
  undefined,
  undefined
)
// 3. 以逗号为分隔符拆分C1，Comma表示用逗号来分隔
Range('C1').TextToColumns(
  undefined,
  xlDelimited,
  xlTextQualifierNone,
  undefined,
  undefined,
  undefined,
  true,
  undefined,
  undefined,
  undefined,
  undefined
)
// 4. 以空格为分隔符拆分D1，Space表示用空格来分隔
Range('D1').TextToColumns(
  undefined,
  xlDelimited,
  xlTextQualifierNone,
  undefined,
  undefined,
  undefined,
  undefined,
  true,
  undefined,
  undefined,
  undefined
)
```

#### 示例二

以@为分隔符拆分 A1。

以#为分隔符分隔 A2。

以*为分隔符分隔 A3。

```js
// 1. 以@为分隔符分隔A1，@不是默认分隔符所以使用Other
Range('A1').TextToColumns(
  undefined,
  xlDelimited,
  xlTextQualifierNone,
  undefined,
  undefined,
  undefined,
  undefined,
  undefined,
  true,
  '@',
  undefined
)
// 2. 以#为分隔符分隔A2，#不是默认分隔符所以使用Other
Range('A2').TextToColumns(
  undefined,
  xlDelimited,
  xlTextQualifierNone,
  undefined,
  undefined,
  undefined,
  undefined,
  undefined,
  true,
  '#',
  undefined
)
// 3. 以*为分隔符分隔A3，*不是默认@分隔符所以使用Other
Range('A3').TextToColumns(
  undefined,
  xlDelimited,
  xlTextQualifierNone,
  undefined,
  undefined,
  undefined,
  undefined,
  undefined,
  true,
  '*',
  undefined
)
```

#### 示例三

以空格和分号分离 A1:A10 的内容，并拆分到 B1。

以逗号、tab 和=为分隔符号分离 B 列的内容并填充到 D 列。

以 tab、逗号、分号和空格为分隔符号拆分 C 列。

```js
// 1. 以空格和分号为分隔符拆分A1:A10，拆分到B1单元格里面
Range('A1:A10').TextToColumns(
  Range('B1'),
  xlDelimited,
  xlTextQualifierNone,
  undefined,
  undefined,
  true,
  undefined,
  true,
  undefined,
  undefined,
  undefined
)
// 2. 以逗号、tab和=为分隔符号拆分B列，拆分到D列里面
Range('B:B').TextToColumns(
  Range('D:D'),
  xlDelimited,
  xlTextQualifierNone,
  undefined,
  true,
  undefined,
  true,
  undefined,
  true,
  '=',
  undefined
)
// 3. 以tab、逗号、分号和空格为分隔符号拆分C列
Range('C:C').TextToColumns(
  undefined,
  xlDelimited,
  xlTextQualifierNone,
  undefined,
  true,
  true,
  true,
  true,
  undefined,
  undefined,
  undefined
)
```

#### 示例四

以逗号对 A1 分列，连续的分隔符要视为单个处理。

以&和空格对 B1:B5 区域进行分列，忽略连续的分隔符。

以!对 C 列拆分数据为多列，连续分隔符单独处理。

```js
// 1. 以逗号为分隔符拆分A1，ConsecutiveDelimiter设为true表示连续分隔符视为一个分隔符
Range('A1').TextToColumns(
  undefined,
  xlDelimited,
  xlTextQualifierNone,
  true,
  undefined,
  undefined,
  true,
  undefined,
  undefined,
  undefined,
  undefined
)
// 2. 以&和tab为分隔符拆分B1:B5，&不是默认分隔符所以使用Other，ConsecutiveDelimiter设为true表示连续分隔符视为一个分隔符
Range('B1:B5').TextToColumns(
  undefined,
  xlDelimited,
  xlTextQualifierNone,
  true,
  undefined,
  undefined,
  undefined,
  true,
  true,
  '&',
  undefined
)
// 3. 以!为分隔符拆分C列，!不是默认分隔符所以使用Other，ConsecutiveDelimiter设为false表示连续分隔符单独处理
Range('C:C').TextToColumns(
  undefined,
  xlDelimited,
  xlTextQualifierNone,
  false,
  undefined,
  undefined,
  undefined,
  undefined,
  true,
  '!',
  undefined
)
```

#### 示例五

以空格为分隔符拆分 A1 内容到 B1，拆分的内容第一列设为常规类型，第二列设为文本类型，第三列设为日期 YMD 格式。

以分号为分隔符拆分 C1 内容到 C3，拆分的内容第二列设为日期 DMY 式，第三列设为日期 MYD 格式，第四列设为日期 DYM,格式，第五列设为日期 YDM 格式。

```js
// 1. 以空格为分隔符拆分A1，拆分到B1, [X, Y] 代表第X列为Y对应的格式，[1, xlGeneralFormat]代表第一列为常规格式，[2, xlTextFormat]代表第二列为文本格式，[3, xlYMDFormat]代表第三列为日期YMD格式，
Range('A1').TextToColumns(
  Range('B1'),
  xlDelimited,
  xlTextQualifierNone,
  undefined,
  undefined,
  undefined,
  undefined,
  true,
  undefined,
  undefined,
  [
    [1, xlGeneralFormat],
    [2, xlTextFormat],
    [3, xlYMDFormat]
  ]
)
// 2. 以分号为分隔符拆分C1，拆分到C3,[2, xlDMYFormat]代表第二列为日期DMY格式，[3, xlMYDFormat]代表第三列为日期MYD格式，[4, xlDYMFormat]代表第四列为日期DYM格式，[5, xlYDMFormat]代表第四列为日期YDM格式，
Range('C1').TextToColumns(
  Range('C3'),
  xlDelimited,
  xlTextQualifierNone,
  undefined,
  undefined,
  true,
  undefined,
  undefined,
  undefined,
  undefined,
  [
    [2, xlDMYFormat],
    [3, xlMYDFormat],
    [4, xlDYMFormat],
    [5, xlYDMFormat]
  ]
)
```
