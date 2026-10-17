---
title: "Sheet"
source: https://airsheet.wps.cn/docs/api/excel/databook/Sheet.html
updated: 1789700916000
hash: BfwEJlCU
---

# Sheet

工作簿（Workbook）中单个数据表(Sheet)对象

Sheet 对象的具体属性和方法请参阅下方的列表。

#### 属性列表

| 属性名 | 数据类型 | 简介 |
| --- | --- | --- |
| [Id](#id) | String | 该数据表的 Id |
| [Name](#name) | String | 该数据表的名称 |
| [Index](#index) | Number | 该数据表在所有表的索引值 |
| [Visible](#visible) | Boolean | 该数据表是否可见 |
| [Type](#type) | String | 该数据表的类型 |
| [Field](#field) | [Field](#field) | 该数据表的字段 |
| [Record](#record) | [Record](#record) | 该数据表的行记录 |

#### 方法列表

| 方法名 | 返回类型 | 简介 |
| --- | --- | --- |
| [Activate()](#activate) | undefined | 切换(激活)数据表 |
| [Move()](#move) | undefined | 移动数据表 |
| [Delete()](#delete) | undefined | 删除数据表 |
| [IsDBSheet()](#isdbsheet) | Boolean | 是否为数据表 |

## Id

获取数据表 Id

#### 数据类型

String - 数据表 Id

#### 示例

```js
const sheet = Application.ActiveSheet
// 打印当前活动数据表的id
console.log(sheet.Id)
```

## Name

设置/获取 数据表名称

#### 数据类型

String - 该数据表在所有数据表的名称

#### 示例

```js
const sheet = Application.ActiveSheet
// 打印当前活动数据表的名称
console.log(sheet.Name) // Sheet2

// 将当前数据表的名称改为 WPS WebOffice
sheet.Name = 'WPS WebOffice'
```

## Index

数据表的 index,即该数据表在所有数据表的索引值

#### 数据类型

String - 该数据表在所有数据表的索引值

#### 示例

```js
const sheet = Application.ActiveSheet
// 打印当前活动数据表的index
console.log(sheet.Index) // 1
```

## Visible

显示/隐藏 数据表

#### 数据类型

Boolean - 数据表是否可见

#### 示例

```js
const sheet = Application.ActiveSheet
// 隐藏数据表
sheet.Visible = false
// 取消数据表隐藏
sheet.Visible = true
```

## Type

数据表类型

#### 数据类型

[Enum.xlEtDataBaseSheet](Enum.md#xlsheettype) - 数据表的类型

#### 示例

```js
const sheet = Application.ActiveSheet
// 打印当前活动数据表的类型
console.log(sheet.Type) //xlEtDataBaseSheet
```

## Field

数据表的字段， 返回一个[Field](Field.md)对象

#### 数据类型

[Field](Field.md)

#### 示例

```js
const sheet = Application.ActiveSheet
// 获取的表所有字段信息
const fields = sheet.Field.GetFields()
```

## Record

数据表的字段， 返回一个[Record](Record.md)对象

#### 数据类型

[Record](Record.md)

#### 示例

```js
const sheet = Application.ActiveSheet
const record = sheet.Record.GetRecord({  RecordId: 'Bz' })
```

## Activate()

激活表

#### 示例

```js
const sheet = Application.Sheets.Item(1)
// 激活第一个表
sheet.Activate()
```

## Move()

移动数据表

#### 参数

**两个参数互斥**

| 属性 | 数据类型 | 默认值 | 必填 | 说明 |
| --- | --- | --- | --- | --- |
| Before | number | null | 否 | 验将放置移动的数据表之前的数据表 ID。如果指定 After ，则不能指定 Before。 |
| After | number | null | 否 | 将放置移动的数据表后的数据表 ID。如果指定 Before ，则不能指定 After |

#### 示例

```js
// 将当前数据表移动到第二个数据表之后
const sheet = Application.ActiveSheet
sheet.Move({
  Before: null,
  After: Application.Sheets(2).Id
})
```

## Delete()

删除数据表

#### 返回值

undefined

#### 示例

```js
// 删除名称为“Sheet2”的数据表
Application.Sheets.Item('Sheet2').Delete()
```

## IsDBSheet()

是否为数据表

#### 返回值

Boolean

#### 示例

```js
// 判断当前活跃表是否为数据表
Application.ActiveSheet.IsDBSheet()
```
