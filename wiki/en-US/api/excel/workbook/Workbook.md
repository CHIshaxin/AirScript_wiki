---
title: "Workbook"
source: https://airsheet.wps.cn/docs/en-US/api/excel/workbook/Workbook.html
updated: 1789700916000
hash: Cds48lBP
---

# Workbook

Workbook: the working area of a file, containing all sheets.

See the lists below for the Workbook object's properties and methods.

#### Properties

| Property | Data type | Description |
|---|---|---|
| [ActiveSheet](#activesheet) | [Sheet](/docs/en-US/api/excel/workbook/Sheet.html) | Active worksheet in the workbook |
| [Sheets](#sheets) | [Sheets](/docs/en-US/api/excel/workbook/Application.html#sheets) | Collection of all worksheet objects |
| [ReadOnly](#readonly) | Boolean | Whether the document is read-only |
| [ReadOnlyComment](#readonlycomment) | Boolean | Whether the document allows comment-only read access |
| [SupportReadOnlyComment](#supportreadonlycomment) | Boolean | Whether the document supports comment-only read access |

#### Methods

| Method | Return type | Description |
|---|---|---|
| [Save()](#save) | String(JSON) | Save changes to the file |
| [GetComments()](#getcomments) | String(JSON) | Get comments for the entire workbook |
| [ExportAsFixedFormat()](#exportasfixedformat) | String(JSON) | Export the whole sheet as PDF or image |

## ActiveSheet

Active worksheet in the workbook.

You can use `Application.ActiveSheet` instead.

#### Data type

[Sheet](/docs/en-US/api/excel/workbook/Sheet.html) - Active worksheet in the active workbook

#### Example

```js
// Both forms below are equivalent
console.log(Application.ActiveWorkbook.ActiveSheet.Name) // Sheet1
console.log(Application.ActiveSheet.Name) // Sheet1
```

## Sheets

Gets all sheets the current file can operate on, as a [Sheets](/docs/en-US/api/excel/workbook/Application.html#sheets) object.

#### Data type

[Sheets](/docs/en-US/api/excel/workbook/Application.html#sheets)

#### Example

```js
// Both forms below are equivalent
let sheets = Application.ActiveWorkbook.Sheets
sheets = Application.Sheets
// Print names of all worksheets
for (let i = 1; i <= sheets.Count; i++) {
    console.log((sheets.Item(i).Name))
}
```

## ReadOnly

Returns whether the document is read-only. This property is read-only.

#### Data type

Boolean - Whether the document is read-only

#### Example

```js
// Print whether the current document is read-only
console.log(Application.ActiveWorkbook.ReadOnly) //false
```

## ReadOnlyComment

Returns whether the document has comment-only read permission. This property is read-only.

#### Data type

Boolean - Whether the document has comment-only read permission

#### Example

```js
// Print whether the current document is comment-only readable
console.log(Application.ActiveWorkbook.ReadOnlyComment) //true
```

## Save()

Saves changes to the file.

#### Return type

String(JSON) - Save result, format as follows:

| Property | Data type | Description |
|---|---|---|
| result | String | Save status |
| size | Number | File size in bytes |
| version | Number | Version |

**Save status values**:

| Status | Description |
|---|---|
| ok | Version saved successfully; viewable in version history |
| nochange | No document changes; no version to save |
| SavedEmptyFile | Saving empty files is not supported. Occurs when the file is empty after kernel save |
| SpaceFull | Storage is full |
| QueneFull | Do not save too frequently. Occurs when the server save queue is full and requests are queued |
| fail | Save failed |

#### Example

```js
// Save changes
let res = Application.ActiveWorkbook.Save()
if (res.result === 'ok') {
  console.log('Success')
} else {
  console.error('Failed:', res.result)
}
```

## GetComments()

Gets comments for the entire workbook.

#### Return type

String(JSON) - Comment collection for the workbook, format as follows:

| Property | Data type | Description |
|---|---|---|
| CellComments | `Array<Object>` | Comment entries |
| PosInfo | `String` | Cell information |
| SheetName | `String` | Sheet name |
| UserIds | `Array<String>` | User id collection |

**Comment entries**:

| Property | Data type | Description |
|---|---|---|
| DateTime | `String` | Timestamp |
| Text | `String` | Comment text |
| Time | `String` | Formatted time |
| UserId | `String` | User id |

#### Example

```js
// Get comments for the entire workbook
let json = Application.ActiveWorkbook.GetComments()
for (let i = 0; i < json.length; i++) {
  let pos = json[i].PosInfo
  let comments = json[i].CellComments
  let commentList = []
  for (let j = 0; j < comments.length; j++) {
    commentList.push(comments[j].Text)
  }
  console.log(pos, ':', commentList)
}
```

## ExportAsFixedFormat()

Exports the whole sheet as PDF or image and returns the export URL.

#### Parameters

| Property | Data type | Default | Required | Description |
|---|---|---|---|---|
| Type | Enum |  | Optional | Export type; see [Enum.XlFixedFormatType](/docs/en-US/api/excel/workbook/Enum.html#xlfixedformattype). Only image and PDF export are supported |

#### Return type

String(JSON) - JSON string containing the export URL

#### Example 1: Export PDF

```js
// Export whole sheet
let json = Application.ActiveWorkbook.ExportAsFixedFormat()
console.log(json.url)
```

#### Example 2: Export image

```js
// Export whole sheet
let json = Application.ActiveWorkbook.ExportAsFixedFormat({
  Type: Application.Enum.XlFixedFormatType.xlTypeIMG
})
console.log(json.url)
```
