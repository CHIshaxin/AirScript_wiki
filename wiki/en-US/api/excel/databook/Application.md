---
title: "Application"
source: https://airsheet.wps.cn/docs/en-US/api/excel/databook/Application.html
updated: 1789700916000
hash: Zq4q7Hrv
---

# Application

The top-level object for document operations. All document-related work is done directly or indirectly through this object.

`Application` is the top-level object for a file. Opening a new file also returns an `Application`.

In scripts, `Application` refers to the top-level object of the current file. There is exactly one.

See the lists below for properties and methods of `Application`.

#### Properties

| Property | Type | Description |
|---|---|---|
| [ActiveSheet](#activesheet) | [Sheet](/docs/en-US/api/excel/databook/Sheet.html) | Active worksheet / data sheet |
| [Sheets](#sheets) | [Sheets](/docs/en-US/api/excel/databook/Application.html#sheets) | All worksheets / data sheets in the file |
| [FileInfo](#fileinfo) | Object | Information about the document |
| [UserInfo](#userinfo) | Object | User information for the document |
| [Enum](#enum) | [Enum](/docs/en-US/api/excel/databook/Enum.html) | All enumeration types |

#### Methods

| Method | Return type | Description |
|---|---|---|
| [Sheets(name)](#sheets-1) | [Sheet](/docs/en-US/api/excel/databook/Sheet.html) | Get the worksheet / data sheet named `name` |

## ActiveSheet

The active worksheet / data sheet. You can switch it with `Sheet.Activate()`. This property returns a [Sheet](/docs/en-US/api/excel/databook/Sheet.html#activesheet) object so you can work with the active sheet.

The script runs in an environment isolated on the server, so the script environment’s `ActiveSheet` may not match the user’s `ActiveSheet`.

Rules:

1. When a script runs, the script environment’s `ActiveSheet` is set to match the user’s current `ActiveSheet`.
2. When the script changes the script environment’s `ActiveSheet`, the user’s `ActiveSheet` does not change.

#### Type

[Sheet](/docs/en-US/api/excel/databook/Sheet.html) – the active worksheet / data sheet

#### Example

```js
console.log(Application.ActiveSheet.Name) // DataSheet2

// Switch to the sheet named DataSheet2
Application.Sheets.Item('DataSheet2').Activate()
console.log(Application.ActiveSheet.Name) // DataSheet2
```

## Sheets

Returns every `Sheet` you can operate on in the current file, as a [Sheets](/docs/en-US/api/excel/databook/Application.html#sheets) object.

#### Type

[Sheets](/docs/en-US/api/excel/databook/Application.html#sheets)

#### Example

```js
// Collection of all worksheets / data sheets (Sheet) in the workbook; these are equivalent
let sheets = Application.ActiveWorkbook.Sheets
sheets = Application.Sheets

// Print names of all worksheets / data sheets
for (let i = 1; i <= sheets.Count; i++) {
  console.log(sheets.Item(i).Name)
}
```

### Sheets.Count

Number of worksheets / data sheets.

#### Type

Number – count of worksheets / data sheets in the workbook

#### Example

```js
// These are equivalent
let sheets = Application.ActiveWorkbook.Sheets
sheets = Application.Sheets

console.log(sheets.Count) // 1
```

### Sheets.DefaultNewSheetName

Default name for a new worksheet.

#### Return type

String – if no name is given when adding a sheet, this name can be used

#### Example

```js
const defaultName = Application.Sheets.DefaultNewSheetName
Application.Sheets.Add(
  null,
  Application.ActiveSheet.Name,
  1,
  Application.Enum.XlSheetType.xlWorksheet,
  defaultName
)
```

### Sheets.Add()

Adds a worksheet. If both `Before` and `After` are provided, `Before` takes precedence.

#### Parameters

| Property | Type | Default | Required | Description |
|---|---|---|---|---|
| Before | String/Number |  | No | Required when `After` is empty: index or name of an existing sheet; the new sheet is inserted before it |
| After | String/Number |  | No | Required when `Before` is empty: index or name of an existing sheet; the new sheet is inserted after it |
| Count | Number | 1 | No | Number of sheets to add. Default is the number of selected sheets |
| Type | Enum |  | No | Sheet type; see [Enum.XlSheetType](/docs/en-US/api/excel/databook/Enum.html#xlsheettype) |
| Name | Name |  | No | Name of the new sheet |

#### Example

```js
Application.Sheets.Add(
  null,
  Application.ActiveSheet.Name,
  1,
  Application.Enum.XlSheetType.xlWorksheet,
  'NewSheet'
)
```

### Sheets.Item()

Select a sheet by name or index.

#### Parameters

| Property | Type | Default | Required | Description |
|---|---|---|---|---|
| index | String/Number |  | Yes | Sheet name or index |

#### Return type

[Sheet](/docs/en-US/api/excel/databook/Sheet.html) – the matching worksheet / data sheet

#### Example

```js
Application.Sheets.Item('Sheet2').Activate()
Application.Sheets.Item(1).Activate()
```

### Sheets.Each()

Iterate all sheets and run a callback.

#### Parameters

| Property | Type | Default | Required | Description |
|---|---|---|---|---|
| callback | Function | null | Yes | Like `Array.prototype.forEach` |

#### Example

```js
Application.Sheets.Each(function (item) {
  console.log(item.Name) // Sheet1 Sheet2
})
```

## FileInfo

Basic information about the current file.

#### Type

Object – file metadata

| Name | Type | Description |
|---|---|---|
| id | string | File ID |
| name | string | File name |
| officeType | string | Document type |
| creator | [CreatorObject](#creatorobject) | Creator info |
| size | number | File size |
| groupId | string | Group ID |
| docType | number | Document type (numeric) |

#### CreatorObject

| Name | Type | Description |
|---|---|---|
| id | string | Creator ID |
| name | string | Creator name |
| avatar_url | string | Avatar URL |
| logined | boolean | Whether logged in |
| attrs | Object | Attributes |
| real_id | string | Real ID |

#### Example

```javascript
console.log(Application.FileInfo)
/*{
 "id": "<open_id>",
 ...
}*/
```

## UserInfo

User information for the current file.

#### Type

Object – user metadata

| Name | Type | Description |
|---|---|---|
| id | string | User ID |
| name | string | User name |

#### Example

```javascript
console.log(Application.UserInfo)
```

## Enum

Enumeration types, exposed on `Application`.

Use `Application.Enum`.

#### Type

[Enum](/docs/en-US/api/excel/databook/Enum.html) – all enumeration types

#### Example

```js
console.log(Application.Enum.XlSheetType)
// {"xlChart":-4109,"xlDialogSheet":-4116,"xlExcel4IntlMacroSheet":4,"xlExcel4MacroSheet":3,"xlWorksheet":-4167}
```

## Sheets()

Used as a function instead of `Sheets.Item()`; returns a [Sheet](/docs/en-US/api/excel/databook/Sheet.html).

#### Parameters

| Name | Type | Required | Description |
|---|---|---|---|
| name | string | Yes | Worksheet / data sheet name |

#### Return type

[Sheet](/docs/en-US/api/excel/databook/Sheet.html) – the sheet object with that name

#### Example

```js
console.log(Application.Sheets.Count) // 1

console.log(Application.Sheets('Sheet2').Range('A1').Text)
console.log(Application.Sheets.Item('Sheet2').Range('A1').Text)
```
