---
title: "Application"
source: https://airsheet.wps.cn/docs/en-US/api/excel/workbook/Application.html
updated: 1789700916000
hash: DBXBxThk
---

# Application

The top-level object for document operations. Working with a document is always done directly or indirectly through this object.

Application is the top-level object for a file; opening a new file also returns Application.

In scripts, `Application` refers to the top-level object of the current file. There is exactly one.

See the lists below for the Application object's properties and methods.

#### Properties

| Property | Data type | Description |
| --- | --- | --- |
| [ActiveSheet](#activesheet) | [Sheet](Sheet.md) | Active worksheet |
| [Sheets](#sheets) | [Sheets](Application.md#sheets) | All worksheets in the current file |
| [ActiveWorkbook](#activeworkbook) | [Workbook](Workbook.md) | Current document |
| [Selection](#selection) | [Range](Range.md) | Current selection |
| [Cells](#cells) | [Range](Range.md) | All cells on the active sheet |
| [Columns](#columns) | [Range](Range.md) | All columns on the active sheet |
| [Rows](#rows) | [Range](Range.md) | All rows on the active sheet |
| [FileInfo](#fileinfo) | Object | Information about the current document |
| [UserInfo](#userinfo) | Object | User information for the current document |
| [Enum](#enum) | [Enum](Enum.md) | All enumeration types |

#### Methods

| Method | Return type | Description |
| --- | --- | --- |
| [Range(address)](#range) | [Range](Range.md) | Gets a region on the current ActiveSheet (address specifies the region) |
| [Sheets(name)](#sheets-1) | [Sheet](Sheet.md) | Gets the worksheet named name |

## ActiveSheet

The active worksheet. You can switch the active sheet with Sheet.Activate(). This property returns a [Sheet](Sheet.md#ActiveSheet) object so you can work with the active worksheet.

The script runs in an environment isolated on the server, so the script environment's ActiveSheet may not match the user's ActiveSheet.

Rules:

1. When a script runs, the script environment's ActiveSheet is switched to match the user's current ActiveSheet.
2. When the script changes the script environment's ActiveSheet through a function, the user's ActiveSheet does not change in sync.

#### Data type

[Sheet](Sheet.md) - Active worksheet

#### Example

```js
console.log(Application.ActiveSheet.Name) // Sheet1

// Switch to the worksheet named Sheet2
Application.Sheets.Item('Sheet2').Activate()
console.log(Application.ActiveSheet.Name) // Sheet2

// Set cell A1 on the active sheet to foo
Application.ActiveSheet.Range('A1').Value = 'foo'
```

## Sheets

Gets all worksheets the current file can operate on, as a [Sheets](Application.md#sheets) object.

#### Data type

[Sheets](Application.md#sheets)

#### Example

```js
// Collection of all worksheets (Sheet) in the workbook (Workbook); both forms below are equivalent
let sheets = Application.ActiveWorkbook.Sheets
sheets = Application.Sheets

// Print names of all worksheets
for (let i = 1; i <= sheets.Count; i++) {
  console.log(sheets.Item(i).Name)
}
```

### Sheets.Count

Number of worksheets.

#### Data type

Number - Number of worksheets in the workbook

#### Example

```js
// Both forms below are equivalent
let sheets = Application.ActiveWorkbook.Sheets
sheets = Application.Sheets

// Print number of worksheets
console.log(sheets.Count) //1
```

### Sheets.DefaultNewSheetName

Default name for a new worksheet.

#### Return type

String - If no name is given when adding a worksheet, this name can be used as the new sheet name.

#### Example

```js
const defaultName = Application.Sheets.DefaultNewSheetName
// Worksheet object
Application.Sheets.Add(
  null,
  Application.ActiveSheet.Name,
  1,
  Application.Enum.XlSheetType.xlWorksheet,
  defaultName
)
```

### Sheets.Add()

Adds a worksheet. If both Before and After are provided, Before takes precedence.

#### Parameters

| Property | Data type | Default | Required | Description |
| --- | --- | --- | --- | --- |
| Before | String/Number |  | No | Required when After is empty: index or name of an existing sheet; the new sheet is placed before this sheet |
| After | String/Number |  | No | Required when Before is empty: index or name of an existing sheet; the new sheet is placed after this sheet |
| Count | Number | 1 | No | Number of sheets to add. Default is the number of selected sheets |
| Type | Enum |  | No | Worksheet type; see [Enum.XlSheetType](Enum.md#xlsheettype) |
| Name | Name |  | No | Worksheet name |

#### Example

```js
// Add worksheet
Application.Sheets.Add(
  null,
  Application.ActiveSheet.Name,
  1,
  Application.Enum.XlSheetType.xlWorksheet,
  'NewSheet'
)
```

### Sheets.Item()

Selects a sheet by name or index.

#### Parameters

| Property | Data type | Default | Required | Description |
| --- | --- | --- | --- | --- |
| index | String/Number |  | Yes | Name or index of the sheet to select |

#### Return type

[Sheet](Sheet.md) - Worksheet with the given name

#### Example

```js
// Activate worksheet named "Sheet2"
Application.Sheets.Item('Sheet2').Activate()

// Activate worksheet at index 1
Application.Sheets.Item(1).Activate()
```

### Sheets.Each()

Iterates all sheets and runs a callback.

#### Parameters

| Property | Data type | Default | Required | Description |
| --- | --- | --- | --- | --- |
| callback | Function | null | Yes | Similar to JS array forEach |

#### Example

```js
// Print names of all worksheets
Application.Sheets.Each(function (item) {
  console.log(item.Name) //Sheet1 Sheet2
})
```

## ActiveWorkbook

The active file. Returns a [Workbook](Workbook.md) object.

#### Data type

[Workbook](Workbook.md) - Active file

#### Example

```js
// Save
Application.ActiveWorkbook.Save()
```

## Selection

The current selection. Returns a [Range](Range.md) object.

#### Data type

[Range](Range.md) - Current selection

#### Example

```js
// Print selection address when user runs the script; here assume user selected A1
console.log(Application.Selection.Address()) // $A$1
```

## Cells

All cells on the current sheet. Returns a [Range](Range.md) object.

#### Data type

[Range](Range.md) - All cells on the current sheet

#### Example

```js
const cells = Application.ActiveSheet.Cells
// Print number of cells in range
console.log(cells.Count) // 17179869184
```

## Columns

All columns on the current sheet. Returns a [Range](Range.md) object.

#### Data type

[Range](Range.md) - All columns on the current sheet

#### Example

```js
const columns = Application.ActiveSheet.Columns
// Print number of columns
console.log(columns.Count) // 16384
```

## Rows

All rows on the current sheet. Returns a [Range](Range.md) object.

#### Data type

[Range](Range.md) - All rows on the current sheet

#### Example

```js
let rows = Application.ActiveSheet.Rows

// Print number of rows
console.log(rows.Count) // 1048576
```

## FileInfo

Returns basic information about the current file.

#### Data type

Object - File information

| Name | Type | Description |
| --- | --- | --- |
| id | string | File ID |
| name | string | File name |
| officeType | string | Document type |
| creator | [CreatorObject](#creatorobject) | Document creator |
| size | number | File size |
| groupId | string | Group ID of the file |
| docType | number | Document type (numeric) |

#### CreatorObject

| Name | Type | Description |
| --- | --- | --- |
| id | string | Creator ID |
| name | string | Creator name |
| avatar_url | string | Creator avatar |
| logined | boolean | Whether logged in |
| attrs | Object | Attribute object |
| real_id | string | Real ID |

#### Example

```javascript
// Print file info
console.log(Application.FileInfo)
/*{
 "id": "<open_id>",
 ...
}*/
```

## UserInfo

Returns user information for the current file.

#### Data type

Object - User information for the current file

| Name | Type | Description |
| --- | --- | --- |
| id | string | User ID |
| name | string | User name |

#### Example

```javascript
// Print user info
console.log(Application.UserInfo)
```

## Enum

Enumeration types stored under Application.

Use via Application.Enum.

#### Data type

[Enum](Enum.md) - All enumeration types

#### Example

```js
// Print worksheet type enum
console.log(Application.Enum.XlSheetType)
//{"xlChart":-4109,"xlDialogSheet":-4116,"xlExcel4IntlMacroSheet":4,"xlExcel4MacroSheet":3,"xlWorksheet":-4167}
```

## Range()

Gets a [Range](Range.md) for a region on the current ActiveSheet.

#### Parameters

| Name | Type | Required | Description |
| --- | --- | --- | --- |
| address | string | Yes | Address: a cell address or range, e.g. "A1" or "A1:A10" |

#### Return type

[Range](Range.md) - Range object for the given address

#### Example

```js
// Both forms below are equivalent: get text of A1 on the active sheet
console.log(Application.Range('A1').Text) // content of A1

console.log(Application.ActiveSheet.Range('A1').Text) // content of A1
```

## Sheets()

Used as a function; equivalent to Sheets.Item(). Returns a [Sheet](Sheet.md) object.

#### Parameters

| Name | Type | Required | Description |
| --- | --- | --- | --- |
| name | string | Yes | Worksheet name |

#### Return type

[Sheet](Sheet.md) - Sheet object with the given name

#### Example

```js
console.log(Application.Sheets.Count) // 1

// Both forms below are equivalent
console.log(Application.Sheets('Sheet2').Range('A1').Text) // content of A1 on Sheet2
console.log(Application.Sheets.Item('Sheet2').Range('A1').Text) // content of A1 on Sheet2
```
