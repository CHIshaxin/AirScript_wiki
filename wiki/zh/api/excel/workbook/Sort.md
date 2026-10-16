---
title: "Sort"
source: https://airsheet.wps.cn/docs/api/excel/workbook/Sort.html
updated: 1789700916000
hash: D5AtpIS8
---

# Sort

排序对象，设置好目标区域后即可进行排序操作。

#### 属性列表

| 属性名 | 数据类型 | 简介 |
|---|---|---|
| [Header](#header) | [XlYesNoGuess](/docs/api/excel/workbook/Enum.html#xlyesnoguess) | 指定第一行是否包含标题信息 |
| [MatchCase](#matchcase) | Boolean | 是否区分大小写 |
| [Orientation](#orientation) | [XlSortOrientation](/docs/api/excel/workbook/Enum.html#xlsortorientation) | 指定排序方向 |
| [Rng](#rng) | [Range](/docs/api/excel/workbook/Range.html) | 返回要执行排序的值的区域 |
| [SortFields](#sortfields) | [SortFields](/docs/api/excel/workbook/SortFields.html) | 该对象代表与 **Sort** 对象关联的排序字段的集合 |
| [SortMethod](#sortmethod) | [XlSortMethod](/docs/api/excel/workbook/Enum.html#xlsortmethod) | 中文排序方法 |

#### 方法列表

| 方法名 | 返回类型 | 简介 |
|---|---|---|
| [Apply()](#apply) | undefined | 根据当前应用的排序状态对区域进行排序 |
| [SetRange()](#setrange) | undefined | 设置排序发生的范围 |

#### 应用示例

以下为您展示排序对象的在工作表内的一些常见应用场景：

示例 1. 对 C 列降序排序并且对 D 列做升序排序

```js
//获取当前表格区域
const range = ActiveSheet.UsedRange
//获取到排序对象
const sort = ActiveSheet.Sort
//获取排序范围
const sortFields = sort.SortFields
//清除之前的范围
sortFields.Clear()
//基于C列降序排序, xlSortOnValues代表按值排序, xlDescending代表降序排序
sortFields.Add(Range('C:C').Item(1, 1), xlSortOnValues, xlDescending)
//基于D列升序排序, xlSortOnValues代表按值排序, xlAscending代表升序排序
sortFields.Add(Range('D:D').Item(1, 1), xlSortOnValues, xlAscending)
//设置是否包含表头参数，xlGuess为自动，xlYes为包含表头，xlNo为不包含表头。默认设置为xlGuess
sort.Header = xlGuess
//设置是否大小写敏感，true为区分大小写，false为不区分大小写，默认设置false
sort.MatchCase = false
//设置中文排序方法，xlPinYin为拼音排序，xlStroke为比划数排序。默认设置为xlPinYin
sort.SortMethod = xlPinYin
//设置排序的方法，xlSortColumns为按列排序，xlSortRows为按行排序，默认设置为xlSortColumns
sort.Orientation = xlSortColumns
//排序前必须设置SetRange
sort.SetRange(range)
//开始排序
sort.Apply()
```

示例 2. 将 G 列单元格颜色为红色的设为顶部，F 列红色的放在末尾

```js
//获取当前表格区域
const range = ActiveSheet.UsedRange
//获取到排序对象
const sort = ActiveSheet.Sort
//获取排序范围
const sortFields = sort.SortFields
//清除之前的范围
sortFields.Clear()
//基于G列把红色放顶部
//增加排序范围。第1个参数为基于排序的单元格区域。第2个参数用来指定排序依据，xlSortOnValues为基于单元格值排序，xlSortOnFontColor为基于字体颜色排序，xlSortOnCellColor基于单元格颜色排序。第3个参数为排序方式，xlAscending为升序，xlDescending为降序
const sortField1 = sortFields.Add(
  Range('G:G').Item(1, 1),
  xlSortOnCellColor,
  xlAscending
)
//设置排序颜色为红色
sortField1.SortOnValue.Color = RGB(255, 0, 0)
//基于F列把红色放末尾
//增加排序范围。第1个参数为基于排序的单元格区域。第2个参数用来指定排序依据，xlSortOnValues为基于单元格值排序，xlSortOnFontColor为基于字体颜色排序，xlSortOnCellColor基于单元格颜色排序。第3个参数为排序方式，xlAscending为升序，xlDescending为降序
const sortField2 = sortFields.Add(
  Range('F:F').Item(1, 1),
  xlSortOnCellColor,
  xlDescending
)
//设置排序颜色为红色
sortField1.SortOnValue.Color = RGB(255, 0, 0)
//设置是否包含表头参数，xlGuess为自动，xlYes为包含表头，xlNo为不包含表头。默认设置为xlGuess
sort.Header = xlGuess
//设置是否大小写敏感，true为区分大小写，flase为不区分大小写，默认设置false
sort.MatchCase = false
//设置中文排序方法，xlPinYin为拼音排序，xlStroke为比划数排序。默认设置为xlPinYin
sort.SortMethod = xlPinYin
//设置排序的方法，xlSortColumns为对列排序，xlSortRows为对行排序，默认设置为xlSortColumns
sort.Orientation = xlSortColumns
//排序前必须设置SetRange
sort.SetRange(range)
//开始排序
sort.Apply()
```

示例 3. 工作表包含 A 列到 H 列，把 B 列中较小的排到前面

```js
// 获取sort对象
const sort = ActiveSheet.Sort

const keyColumn = 'B'
const keyBeginRow = ActiveSheet.UsedRange.Row
const keyEndRow =
  ActiveSheet.UsedRange.Row + ActiveSheet.UsedRange.Rows.Count - 1

const sortFields = sort.SortFields
const sortKeyRange = ActiveSheet.Range(
  keyColumn + keyBeginRow + ':' + keyColumn + keyEndRow
)
sortFields.Clear()
//xlSortOnValues代表按值排序, xlAscending代表升序排序
sortFields.Add(sortKeyRange, xlSortOnValues, xlAscending)

//选中排序区域
const applyBeginRow = ActiveSheet.UsedRange.Row
const applyEndRow =
  ActiveSheet.UsedRange.Row + ActiveSheet.UsedRange.Rows.Count - 1
const applyRange = ActiveSheet.Range(`A${applyBeginRow}:H${applyEndRow}`)
// 排序前必须设置范围
sort.SetRange(applyRange)
//应用排序
sort.Apply()
```

示例 4. 工作表包含 A 列到 H 列，按笔画从大到小排列 A 列

```js
//获取 sort 对象
const sort = ActiveSheet.Sort

//选中关键区域
const keyColumn = 'A'
const keyBeginRow = ActiveSheet.UsedRange.Row
const keyEndRow =
  ActiveSheet.UsedRange.Row + ActiveSheet.UsedRange.Rows.Count - 1
const keyRangeStr = keyColumn + keyBeginRow + ':' + keyColumn + keyEndRow

//进行数据降序排序
const sortFields = sort.SortFields
sortFields.Clear()
sortFields.Add(
  ActiveSheet.Range(keyRangeStr),
  xlSortOnValues,
  xlDescending,
  '',
  undefined
)
sort.header = xlYes
sort.Orientation = xlSortColumns
//设置为按笔画排列
sort.SortMethod = xlStroke

//设置排序区域
const beginRow = ActiveSheet.UsedRange.Row
const endRow = ActiveSheet.UsedRange.Row + ActiveSheet.UsedRange.Rows.Count - 1

// 排序前必须设置范围
sort.SetRange(ActiveSheet.Range(`A${beginRow}:H${endRow}`))
//应用排序
sort.Apply()
```

示例 5. 工作表包含 A 列到 H 列，按 B 列单元格颜色排列，无色的在顶端

```js
//选中关键区域
const keyColumn = 'B'
const keyBeginRow = ActiveSheet.UsedRange.Row
const keyEndRow =
  ActiveSheet.UsedRange.Row + ActiveSheet.UsedRange.Rows.Count - 1
const keyRangeStr = keyColumn + keyBeginRow + ':' + keyColumn + keyEndRow
let keyRange = ActiveSheet.Range(keyRangeStr)

//单元格颜色进行升序排列
let sort = ActiveSheet.Sort
let sortFields = sort.SortFields
sortFields.Clear()
sortFields.Add(keyRange, xlSortOnCellColor, xlDescending, '', undefined)
sort.header = xlYes
sort.Orientation = xlSortColumns
sort.SortMethod = xlPinYin

//设置排序区域
const beginRow = ActiveSheet.UsedRange.Row
const endRow = ActiveSheet.UsedRange.Row + ActiveSheet.UsedRange.Rows.Count - 1

// 排序前必须设置范围
sort.SetRange(ActiveSheet.Range(`A${beginRow}:H${endRow}`))
//应用排序
sort.Apply()
```

## Header

指定第一行是否包含标题信息，可读/写 [XlYesNoGuess](/docs/api/excel/workbook/Enum.html#xlyesnoguess)。

默认值为 xlNo，如果希望 Excel 确定标题，可以指定 xlGuess。

#### 数据类型

[XlYesNoGuess](/docs/api/excel/workbook/Enum.html#xlyesnoguess) - 第一行是否包含标题

## MatchCase

设置为 `true` 以执行区分大小写的排序，或设置为 `false` 以执行不区分大小写的排序。可读/写。

#### 数据类型

Boolean - 是否区分大小写

## Orientation

指定排序方向，可读/写 [XlSortOrientation](/docs/api/excel/workbook/Enum.html#xlsortorientation)。

#### 数据类型

[XlSortOrientation](/docs/api/excel/workbook/Enum.html#xlsortorientation) - 排序方向

## Rng

返回要执行排序的值的区域，此为只读属性。

#### 数据类型

[Range](/docs/api/excel/workbook/Range.html) - 排序区域对象

## SortFields

代表与 **Sort** 对象关联的排序字段的集合，此为只读属性。

### Count

返回集合中对象的数目，只读。

#### 数据类型

Number - 对象数目

### Add()

创建新的排序字段，并返回一个 SortField 对象

#### 参数

| 属性 | 数据类型 | 默认值 | 必填 | 说明 |
|---|---|---|---|---|
| Key | [Range](/docs/api/excel/workbook/Range.html) | null | 是 | 指定用于排序的键值 |
| SortOn | [XlSortOn](/docs/api/excel/workbook/Enum.html#xlsorton) | null | 否 | 要进行排序的字段 |
| Order | [XlSortOrder](/docs/api/excel/workbook/Enum.html#xlsortorder) | null | 否 | 指定排序次序 |
| CustomOrder | Variant | null | 否 | 指定是否应使用自定义排序次序 |
| DataOption | [XlSortDataOption](/docs/api/excel/workbook/Enum.html#xlsortdataoption) | null | 否 | 指定数据选项 |

#### 返回值

[SortField](/docs/api/excel/workbook/SortField.html) - 代表与 **Sort** 对象关联的排序字段对象

### Clear()

清除所有 SortFields 对象，在开始添加排序字段时，最好先调用一次此方法

#### 返回值

undefined

### Item()

返回一个 [SortField](/docs/api/excel/workbook/SortField.html) 对象，此为只读属性

#### 参数

| 属性 | 数据类型 | 默认值 | 必填 | 说明 |
|---|---|---|---|---|
| Index | Number / String | null | 是 | 指定用于索引，默认从 1 开始 |

#### 返回值

[SortField](/docs/api/excel/workbook/SortField.html) - 代表与 **Sort** 对象关联的排序字段对象

## SortMethod

指定中文排序方法，可读/写 [XlSortMethod](/docs/api/excel/workbook/Enum.html#xlsortmethod)。

#### 数据类型

[XlSortMethod](/docs/api/excel/workbook/Enum.html#xlsortmethod) - 中文排序方法

## Apply()

根据当前应用的排序状态对区域进行排序

注意

应用排序规则前必须先进行[SetRange](#setrange)，否则排序不会生效。

#### 返回值

undefined

## SetRange()

设置排序发生的范围

#### 参数

| 属性 | 数据类型 | 默认值 | 必填 | 说明 |
|---|---|---|---|---|
| Range | [Range](/docs/api/excel/workbook/Range.html) | null | 是 | 指定 Sort 对象所表示的排序所依据的范围 |

#### 返回值

undefined
