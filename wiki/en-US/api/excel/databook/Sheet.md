---
title: "Sheet"
source: https://airsheet.wps.cn/docs/en-US/api/excel/databook/Sheet.html
updated: 1789700916000
hash: 80UOjWqV
---

# Sheet

A single data sheet (`Sheet`) object inside a workbook (`Workbook`).

See the lists below for properties and methods.

#### Properties

| Name | Type | Description |
| --- | --- | --- |
| [Id](#id) | String | Sheet ID |
| [Name](#name) | String | Sheet name |
| [Index](#index) | Number | Index among all sheets |
| [Visible](#visible) | Boolean | Whether the sheet is visible |
| [Type](#type) | String | Sheet type |
| [Field](#field) | [Field](#field) | Fields on this sheet |
| [Record](#record) | [Record](#record) | Row records on this sheet |

#### Methods

| Name | Return type | Description |
| --- | --- | --- |
| [Activate()](#activate) | undefined | Activate the sheet |
| [Move()](#move) | undefined | Move the sheet |
| [Delete()](#delete) | undefined | Delete the sheet |
| [IsDBSheet()](#isdbsheet) | Boolean | Whether this is a data sheet |

## Id

Gets the sheet ID.

#### Type

String – sheet ID

#### Example

```js
const sheet = Application.ActiveSheet
console.log(sheet.Id)
```

## Name

Get or set the sheet name.

#### Type

String – name among all sheets

#### Example

```js
const sheet = Application.ActiveSheet
console.log(sheet.Name) // Sheet2
sheet.Name = 'WPS WebOffice'
```

## Index

The sheet’s index among all sheets.

#### Type

String – index among all sheets (as in original API)

#### Example

```js
const sheet = Application.ActiveSheet
console.log(sheet.Index) // 1
```

## Visible

Show or hide the sheet.

#### Type

Boolean – visibility

#### Example

```js
const sheet = Application.ActiveSheet
sheet.Visible = false
sheet.Visible = true
```

## Type

Sheet type.

#### Type

[Enum.xlEtDataBaseSheet](Enum.md#xlsheettype) – sheet type

#### Example

```js
const sheet = Application.ActiveSheet
console.log(sheet.Type) // xlEtDataBaseSheet
```

## Field

Fields on the data sheet; returns a [Field](Field.md) object.

#### Type

[Field](Field.md)

#### Example

```js
const sheet = Application.ActiveSheet
const fields = sheet.Field.GetFields()
```

## Record

Row records on the data sheet; returns a [Record](Record.md) object.

#### Type

[Record](Record.md)

#### Example

```js
const sheet = Application.ActiveSheet
const record = sheet.Record.GetRecord({ RecordId: 'Bz' })
```

## Activate()

Activates the sheet.

#### Example

```js
const sheet = Application.Sheets.Item(1)
sheet.Activate()
```

## Move()

Moves the sheet.

#### Parameters

**`Before` and `After` are mutually exclusive.**

| Property | Type | Default | Required | Description |
| --- | --- | --- | --- | --- |
| Before | number | null | No | ID of the sheet before which this sheet is placed. Cannot be used together with `After`. |
| After | number | null | No | ID of the sheet after which this sheet is placed. Cannot be used together with `Before`. |

#### Example

```js
const sheet = Application.ActiveSheet
sheet.Move({
  Before: null,
  After: Application.Sheets(2).Id
})
```

## Delete()

Deletes the sheet.

#### Return value

undefined

#### Example

```js
Application.Sheets.Item('Sheet2').Delete()
```

## IsDBSheet()

Whether the sheet is a data sheet.

#### Return value

Boolean

#### Example

```js
Application.ActiveSheet.IsDBSheet()
```
