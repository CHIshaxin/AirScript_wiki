---
title: "Application"
source: https://airsheet.wps.cn/docs/api/excel/workbook/Application.html
updated: 1789700916000
hash: We3On_J1
---

# Application

文档操作的顶级对象，对文档进行相关操作，都是间接或直接操作该对象。

Application 是一个文件的顶级对象，新打开一个文件返回的也是 Application。

而在脚本中的`Application`则是指当前文件的顶级对象，有且只有一个。

Application 对象的具体属性和方法请参阅下方的列表。

#### 属性列表

| 属性 | 数据类型 | 简介 |
| --- | --- | --- |
| [ActiveSheet](#activesheet) | [Sheet](Sheet.md) | 当前的活动工作表 |
| [Sheets](#sheets) | [Sheets](Application.md#sheets) | 当前文件的所有工作表 |
| [ActiveWorkbook](#activeworkbook) | [Workbook](Workbook.md) | 当前的文档 |
| [Selection](#selection) | [Range](Range.md) | 当前的选区对象 |
| [Cells](#cells) | [Range](Range.md) | 当前工作表所有单元格 |
| [Columns](#columns) | [Range](Range.md) | 当前工作表所有列 |
| [Rows](#rows) | [Range](Range.md) | 当前工作表所有行 |
| [FileInfo](#fileinfo) | Object | 当前文档的信息 |
| [UserInfo](#userinfo) | Object | 当前文档的用户信息 |
| [Enum](#enum) | [Enum](Enum.md) | 所有的枚举类型 |

#### 方法列表

| 方法 | 返回类型 | 简介 |
| --- | --- | --- |
| [Range(address)](#range) | [Range](Range.md) | 获取当前 ActiveSheet 的某个区域(address 指定) |
| [Sheets(name)](#sheets-1) | [Sheet](Sheet.md) | 获取名称为 name 的工作表 |

## ActiveSheet

当前活动工作表，可以通过 Sheet.Activate()来切换活动工作表。该属性返回[Sheet](Sheet.md#ActiveSheet)对象,能利用该属性操作当前活动工作表。

运行脚本的环境是独立在服务器的，因此脚本运行环境的 ActiveSheet 与用户环境的 ActiveSheet 不一定相同。

具体规则是：

1.运行脚本时会把脚本运行环境的 ActiveSheet 切换为用户环境当前的 ActiveSheet。

2.当脚本通过函数切换脚本运行环境的 ActiveSheet 时，用户环境的 ActiveSheet 不会同步切换。

#### 数据类型

[Sheet](Sheet.md) - 当前活动工作表

#### 示例

```js
console.log(Application.ActiveSheet.Name) // Sheet1

// 切换到名称为Sheet2的工作表
Application.Sheets.Item('Sheet2').Activate()
console.log(Application.ActiveSheet.Name) // Sheet2

// 将活动工作表的A1单元格赋值foo
Application.ActiveSheet.Range('A1').Value = 'foo'
```

## Sheets

获取当前文件能操作的所有 Sheet，返回一个[Sheets](Application.md#sheets)对象。

#### 数据类型

[Sheets](Application.md#sheets)

#### 示例

```js
// 工作簿（Workbook）中所有工作表（Sheet）的集合,下面两种写法是一样的
let sheets = Application.ActiveWorkbook.Sheets
sheets = Application.Sheets

// 打印所有工作表的名称
for (let i = 1; i <= sheets.Count; i++) {
  console.log(sheets.Item(i).Name)
}
```

### Sheets.Count

工作表数量

#### 数据类型

Number - 对应工作簿的工作表数量

#### 示例

```js
// 下面两种写法是一样的
let sheets = Application.ActiveWorkbook.Sheets
sheets = Application.Sheets

// 打印所有工作表的名称
console.log(sheets.Count) //1
```

### Sheets.DefaultNewSheetName

默认新工作表名

#### 返回类型

String - 新建工作表时若没有指定名称，可用这个名称作为新建工作表名称

#### 示例

```js
const defaultName = Application.Sheets.DefaultNewSheetName
// 工作表对象
Application.Sheets.Add(
  null,
  Application.ActiveSheet.Name,
  1,
  Application.Enum.XlSheetType.xlWorksheet,
  defaultName
)
```

### Sheets.Add()

新增工作表，如果 Before 和 After 都存在，以 Before 为准

#### 参数

| 属性 | 数据类型 | 默认值 | 必填 | 说明 |
| --- | --- | --- | --- | --- |
| Before | String/Number |  | 否 | After 空时，必填，为当前已有单元格的 index 或者名称，新建的工作表将置于此工作表之前 |
| After | String/Number |  | 否 | Before 空时，必填，为当前已有单元格的 index 或者名称，新建的工作表将置于此工作表之后 |
| Count | Number | 1 | 否 | 要添加的工作表数。默认值为选定工作表的数量 |
| Type | Enum |  | 否 | 指定工作表类型，详细可见 [Enum.XlSheetType](Enum.md#xlsheettype) |
| Name | Name |  | 否 | 指定工作表名称 |

#### 示例

```js
// 添加工作表
Application.Sheets.Add(
  null,
  Application.ActiveSheet.Name,
  1,
  Application.Enum.XlSheetType.xlWorksheet,
  '新工作表'
)
```

### Sheets.Item()

根据名称或索引选择 Sheet

#### 参数

| 属性 | 数据类型 | 默认值 | 必填 | 说明 |
| --- | --- | --- | --- | --- |
| index | String/Number |  | 是 | 所选的 sheet 的名称/索引 |

#### 返回类型

[Sheet](Sheet.md) - 对应名称的工作表

#### 示例

```js
// 切换名称为"Sheet2"的工作表
Application.Sheets.Item('Sheet2').Activate()

// 切换索引为1的工作表
Application.Sheets.Item(1).Activate()
```

### Sheets.Each()

遍历所有 sheet 并执行回调函数

#### 参数

| 属性 | 数据类型 | 默认值 | 必填 | 说明 |
| --- | --- | --- | --- | --- |
| callback | Function | null | 是 | 类似 JS 数组的 forEach |

#### 示例

```js
// 打印所有工作表的名称
Application.Sheets.Each(function (item) {
  console.log(item.Name) //Sheet1 Sheet2
})
```

## ActiveWorkbook

当前活动的文件。返回一个[Workbook](Workbook.md)对象。

#### 数据类型

[Workbook](Workbook.md) - 当前活动文件

#### 示例

```js
//保存
Application.ActiveWorkbook.Save()
```

## Selection

当前的选区对象。返回一个 [Range](Range.md) 对象。

#### 数据类型

[Range](Range.md) - 当前的选区对象

#### 示例

```js
// 打印用户运行脚本时页面的选区地址,此处假设用户选择了A1
console.log(Application.Selection.Address()) // $A$1
```

## Cells

当前 Sheet 的所有单元格。返回一个[Range](Range.md)对象。

#### 数据类型

[Range](Range.md) - 当前 Sheet 的所有单元格

#### 示例

```js
const cells = Application.ActiveSheet.Cells
// 打印range里单元格的数量
console.log(cells.Count) // 17179869184
```

## Columns

当前 Sheet 的所有列。返回一个[Range](Range.md)对象。

#### 数据类型

[Range](Range.md) - 当前 Sheet 的所有列

#### 示例

```js
const columns = Application.ActiveSheet.Columns
// 打印列的数量
console.log(columns.Count) // 16384
```

## Rows

当前 Sheet 的所有行。返回一个[Range](Range.md)对象。

#### 数据类型

[Range](Range.md) - 当前 Sheet 的所有行

#### 示例

```js
let rows = Application.ActiveSheet.Rows

// 打印行的数量
console.log(rows.Count) // 1048576
```

## FileInfo

返回当前文件的基本信息。

#### 数据类型

Object - 当前文件的信息

| 名称 | 类型 | 说明 |
| --- | --- | --- |
| id | string | 文件 ID |
| name | string | 文件名 |
| officeType | string | 文档类型 |
| creator | [CreatorObject](#creatorobject) | 文档创建者信息 |
| size | number | 文件大小 |
| groupId | string | 文件的群组 ID |
| docType | number | 文档类型（数字形式） |

#### CreatorObject 对象信息

| 名称 | 类型 | 说明 |
| --- | --- | --- |
| id | string | 创建者 ID |
| name | string | 创建者名称 |
| avatar_url | string | 创建者头像 |
| logined | boolean | 是否已登录 |
| attrs | Object | 属性对象 |
| real_id | string | 真实 ID |

#### 示例

```javascript
// 打印文件信息
console.log(Application.FileInfo)
/*{
 "id": "<open_id>",
 ...
}*/
```

## UserInfo

返回当前文件的用户信息。

#### 数据类型

Object- 当前文件的用户信息

| 名称 | 类型 | 说明 |
| --- | --- | --- |
| id | string | 用户 ID |
| name | string | 用户名称 |

#### 示例

```javascript
// 打印用户信息
console.log(Application.UserInfo)
```

## Enum

枚举类型，存放在 Application 下。

可以通过 Application.Enum 使用

#### 数据类型

[Enum](Enum.md) - 所有的枚举类型

#### 示例

```js
// 打印工作表的类型枚举
console.log(Application.Enum.XlSheetType)
//{"xlChart":-4109,"xlDialogSheet":-4116,"xlExcel4IntlMacroSheet":4,"xlExcel4MacroSheet":3,"xlWorksheet":-4167}
```

## Range()

获取当前 ActiveSheet 的某个区域的[Range](Range.md)对象。

#### 参数

| 名称 | 类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| address | string | 是 | 选择的地址，可以是单元格地址或者选区范围范围地址，例"A1"或"A1:A10" |

#### 返回类型

[Range](Range.md) - 给定地址区域的 Range 对象

#### 示例

```js
// 以下两种写法的效果是一样的,都是获取当前激活的Sheet的A1单元格的内容
console.log(Application.Range('A1').Text) // A1单元格的内容

console.log(Application.ActiveSheet.Range('A1').Text) // A1单元格的内容
```

## Sheets()

作为函数使用，代替 Sheets.Item()，返回一个[Sheet](Sheet.md)对象。

#### 参数

| 名称 | 类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| name | string | 是 | 工作表的名称 |

#### 返回类型

[Sheet](Sheet.md) - 对应名称的工作表 Sheet 对象

#### 示例

```js
console.log(Application.Sheets.Count) // 1

// 以下两种写法效果是一样的
console.log(Application.Sheets('Sheet2').Range('A1').Text) // Sheet2的A1单元格的内容
console.log(Application.Sheets.Item('Sheet2').Range('A1').Text) // Sheet2的A1单元格的内容
```
