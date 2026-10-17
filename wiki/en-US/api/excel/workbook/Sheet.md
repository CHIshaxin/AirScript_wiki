---
title: "Sheet"
source: https://airsheet.wps.cn/docs/en-US/api/excel/workbook/Sheet.html
updated: 1789700916000
hash: D-unBUqt
---

# Sheet

A single worksheet (Sheet) object in a workbook (Workbook).

See the lists below for the Sheet object's properties and methods.

#### Properties

| Property | Data type | Description |
| --- | --- | --- |
| [Id](#id) | String | Id of this worksheet |
| [Name](#name) | String | Name of this worksheet |
| [Index](#index) | Number | Index of this sheet among all sheets |
| [Cells](#cells) | [Range](Range.md) | All cells on this sheet |
| [Columns](#columns) | [Range](Range.md) | All columns on this sheet |
| [Rows](#rows) | [Range](Range.md) | All rows on this sheet |
| [UsedRange](#usedrange) | [Range](Range.md) | Used range on this sheet |
| [Visible](#visible) | Boolean | Whether this sheet is visible |
| [Type](#type) | String | Type of this worksheet |
| [Hyperlinks](#hyperlinks) | [Hyperlinks](#hyperlinks) | All hyperlinks on this sheet |
| [Shapes](#shapes) | [Shapes](#shapes) | All shapes on this sheet |
| [Sort](#sort) | [Sort](Sort.md) | Sort object on this sheet |

#### Methods

| Method | Return type | Description |
| --- | --- | --- |
| [Range()](#range) | [Range](Range.md) | A cell or cell range |
| [Cells()](#cells-1) | [Range](Range.md) | A cell on this sheet |
| [Activate()](#activate) | undefined | Activate worksheet |
| [Move()](#move) | undefined | Move worksheet |
| [Delete()](#delete) | undefined | Delete worksheet |

## Id

Gets the worksheet id

#### Data type

String - Worksheet id

#### Example

```js
const sheet = Application.ActiveSheet
// Print id of active worksheet
console.log(sheet.Id)
```

## Name

Gets or sets the worksheet name

#### Data type

String - Name of this sheet among all sheets

#### Example

```js
const sheet = Application.ActiveSheet
// Print name of active worksheet
console.log(sheet.Name) // Sheet2

// Rename current sheet to WPS WebOffice
sheet.Name = 'WPS WebOffice'
```

## Index

Sheet index: position among all worksheets

#### Data type

String - Index among all worksheets

#### Example

```js
const sheet = Application.ActiveSheet
// Print index of active worksheet
console.log(sheet.Index) // 1
```

## Cells

Collection of all cells on the sheet as a Range (use Range properties and methods).

#### Data type

[Range](Range.md) - All cells on the sheet

#### Example

```js
const sheet = Application.ActiveSheet
// Print full cell address range of active sheet
console.log(sheet.Cells.Address()) // $A$1:$XFD$1048576
```

## Columns

All columns on the sheet as a Range; see [Range](Range.md).

#### Data type

[Range](Range.md) - All columns on the sheet

#### Example

```js
// Print number of columns on this sheet
const sheet = Application.ActiveSheet
console.log(sheet.Columns.Count) //16384
```

## Rows

Rows on the sheet as a Range (use Range properties and methods).

#### Data type

[Range](Range.md) - All columns on the sheet

#### Example

```js
const sheet = Application.ActiveSheet
console.log(sheet.Rows.Count) // 1048576
```

## UsedRange

Used range on the sheet (actually used area) as a Range (use Range properties and methods).

#### Data type

[Range](Range.md) - A cell or range on the sheet

#### Example

```js
const UsedRange = Application.ActiveSheet.UsedRange
// Print used range bounds; here assume 10x10 used area
console.log(
  UsedRange.Row,
  UsedRange.RowEnd,
  UsedRange.Column,
  UsedRange.ColumnEnd
) // 1,10,1,10
```

## Visible

Show or hide the worksheet

#### Data type

Boolean - Whether the sheet is visible

#### Example

```js
const sheet = Application.ActiveSheet
// Hide sheet
sheet.Visible = false
// Show sheet again
sheet.Visible = true
```

## Type

Worksheet type

#### Data type

[Enum.XlSheetType](Enum.md#xlsheettype) - Worksheet type

#### Example

```js
const sheet = Application.ActiveSheet
// Print type of active worksheet
console.log(sheet.Type) //xlWorksheet
```

## AutoFilter

AutoFilter for the current sheet: all column filters and methods to apply or clear filters.

#### Data type

[AutoFilter](AutoFilter.md) - AutoFilter object

#### Example

```js
// Get AutoFilter
const autoFilter = Application.ActiveSheet.AutoFilter
// Get filter collection
const filters = autoFilter.Filters
// Count filters on current sheet
const count = filters.Count
console.log(count)
```

## Hyperlinks

Collection of all hyperlinks on the sheet.

#### Data type

[Hyperlinks](#hyperlinks) collection - all hyperlink objects

#### Example

```js
// Print number of hyperlinks
const hyperlinks = Application.ActiveSheet.Hyperlinks
console.log(hyperlinks.Count) // 3
```

### Hyperlinks.Count

Number of hyperlinks in the collection

#### Data type

Number - Number of hyperlinks in the collection

#### Example

```js
// Print number of hyperlinks
const hyperlinks = Application.ActiveSheet.Hyperlinks
console.log(hyperlinks.Count) // 3
```

### Hyperlinks.Item()

Gets a hyperlink object

#### Parameters

| Property | Data type | Default | Required | Description |
| --- | --- | --- | --- | --- |
| Index | Number |  | Yes | Starts at 1 |

#### Return type

[Hyperlink](Hyperlink.md) - Hyperlink object

#### Example

```js
const hyperlinks = Application.ActiveSheet.Hyperlinks
if (hyperlinks.Count <= 0) {
  throw new Error('No hyperlinks in document')
}
// Print hyperlink display text and address
console.log(hyperlinks.Item(1).TextToDisplay) // Hyperlink display text
console.log(hyperlinks.Item(1).Address) // Hyperlink target URL
```

## Shapes

Collection of all Shape objects on the current sheet.

#### Data type

[Shapes](#shapes) collection - all Shape objects on this sheet

#### Example

```js
const sheet = Application.ActiveSheet
const shapes = sheet.Shapes
const chartEnum = Application.Enum.XlChartType.xlColumnClustered // clustered column
// Add 300x300 clustered column chart
shapes.AddChart2(340, chartEnum, 0, 0, 300, 300)
```

### Shapes.GetActiveShapeImg()

Gets image data for the active cell shape

#### Return type

String - Original image download URL

#### Example

```js
// If A1 has a picture
Application.Range('A1').Select()
const imgUrl = Application.ActiveSheet.Shapes.GetActiveShapeImg()
console.log(imgUrl) // https://imageUrl or undefined if none
```

### Shapes.Item(Index)

Represents a drawing-layer object such as a freeform, polygon, OLE object, or picture.

#### Parameters

| Name | Data type | Default | Optional | Description |
| --- | --- | --- | --- | --- |
| index | Number |  | No | Object index |

#### Return type

[Shape](Shape.md) - One drawing-layer object (freeform, polygon, OLE, picture)

#### Example

```js
const shapes = Application.ActiveSheet.Shapes
// Print first shape ID
console.log(shapes.Item(1).ID)
```

### Shapes.Count

Number of shapes

#### Data type

Number - Number of shape objects on the sheet

#### Example

```js
const sheet = Application.ActiveSheet
const shapes = sheet.Shapes
// Print current shape count
console.log(shapes.Count)
```

### Shapes.AddChart2()

Adds a shape

#### Parameters

| Property | Data type | Default | Required | Description |
| --- | --- | --- | --- | --- |
| Style | String |  | No | Chart style |
| XlChartType | Enum |  | No | Chart type; see [Enum.XlChartType](Enum.md#xlcharttype) |
| Left | Number |  | No | Left offset of new chart in px |
| Top | Number |  | No | Top offset of new chart in px |
| Width | Number |  | No | Width of new chart in px |
| Height | Number |  | No | Height of new chart in px |

#### Return type

[Shape](Shape.md) - Added shape

#### Example

```js
const chartEnum = Application.Enum.XlChartType.xlColumnClustered // clustered column enum
const sheet = Application.ActiveSheet
const shapes = sheet.Shapes
// Add 300x300 clustered column chart
shapes.AddChart2(340, chartEnum, 0, 0, 300, 300)
```

### Shapes.AddPicture()

Inserts a floating picture into the sheet

#### Parameters

| Property | Data type | Default | Required | Description |
| --- | --- | --- | --- | --- |
| FileName | String |  | Yes | Image to insert (Base64 data URL or image URL) |
| LinkToFile | Enum |  | Yes | Not supported; use `0` |
| SaveWithDocument | Enum |  | Yes | Not supported; use `0` |
| Left | Number |  | Yes | Distance from left edge of sheet to left edge of image |
| Top | Number |  | Yes | Distance from top edge of sheet to top edge of image |
| Width | Number |  | No | Image width |
| Height | Number |  | No | Image height |

#### Example

```js
// Get shapes collection
const shapes = ActiveSheet.Shapes
// Insert floating picture
shapes.AddPicture({
  FileName:
    'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAACAAAAAgCAMAAABEpIrGAAACH1BMVEUAAAA9kP8mpv9Fv/8QT94TXeQAfv8trP0xr/0VUdoAcv8LgPs1sf4AgP4QUuURTuEVl/ksq/wAg/84tf4NUt80sf0bnPsPUuAAYfsAgP4UTtYAY/8GXPQRkv1Jxv8AZP86tv9Bvv9Gv/9Atv8KiPUOVesAZ+wbS8wUl/kjo/sAYv8PVekjovxBu/8AY/8trPxCu/8AW/AAcvAxsP0npvsYTdElo/oQT9o+uP8AYv8KVOIHdugAYfwAfPwbfe49uf4XmfkcdOopo/oAYv86tf4AYf8TT9cZaeYVTdUJVOUcnvsAYv8Agv9Iwv8AYv9Jwf8bSs0gofoAY/8Ag/9Gv/8en/oAYf9Fv/8AWuQZS80AdO5BvP8AYv9DvP8Umfo9uv8LVOUAgf8amfk/uf8AY/8Ag/8Ahf8TUdgVT9UNWOknov8trPwxr/wmpvspqfs1sv0io/s5tf0AgP8bnfoMV+4AYPoAYv8AYf4AYfwAe/sen/oAd/gAXvYAff0LW/MLWfEOUN0PT9oSTtZAu/49uP0qqfweoPoXmvkAdPUAXPIAcfEKUuIWTNIXS88KXPYAWusOVeoPVOkPUuYQUeQQT+IMUeAUTdMZSs09uP4Aa+kAWOgAaeQIVOQIU+QAV+MRTd8STd0RTtgSl/kAW+8NWO8Abu0AbuwBZNsAg/8AYPwAc/MIVecAZuEAVeAAWt8AZN0JgvAIfOsJV+sTVuIAU9qil9AQAAAAa3RSTlMAAwb++A3+bkkkGxL8+Pf36tnTubCYjol4dmVjRjEvLysZFA7+/Pv6+fj39/b19fPw8O/u7Ovq6Ofn5+Xj4uDa2trZ1dPS0s/Ny7+9vLi4trapoqGgoJ2XkI+KgoJycGRhX1BNS0pHQj06IcCB3jkAAAIBSURBVDjLdZCHctpAFEWfCMU1ce8tTu+99957770nJFGwkE1sgyMTAginYbkAIQSD494+0G+ltTAacWZnNLt7dO/Og3Tqzl2og8wsvroAudEI+iy5s+6LzMZ7Bp1r5sFOQegRenAJwoEnoCXr6Mc0ihvSrhuK32vJK21Mva00760OWyz0KZYNbzJQmKXUF1o7rZ24rGngweETIGOwbH2nw/ZDUaM6grJVHzSs3R+NoqDy+lKLuwWX2618r03/QRalhNzbewo+qZyuZX4TVMFQubm3f3xy5VeZg49xsL8IC+mUH+1zeLyJ3pHkRCwWK7hvIGc/CYrwvIhtbXZ4hgYwJLm0LFf56QdhDRAq8wdZMypD3kT/SBJo6nfCMmWTc2WQlUO8GDIn2GSAUnuWGERJpIS4LR4HlWdFNITua9pl5s3h5vJ81tzscAAhuyTc3kZQ51CxqasvMmDGENyZyo1BX2iqLRAI0L7qvU2fv6Ex5mVbgana4XIGfb7QsI0KNcc4O4/Gv67I6Jjn5cmOvy4nGuFwaFipqFjN2YlBQkZn/N0dxAhiSfjUKzqHyxwq/0lIX0T0d6OAGb7dD0Gl/gzH8UqN6JcNp7HcBPN5esTO8U2oSCIpcZVkgwamepedxxpJROP4C9DBdHc9vlWSxG1VDOiTc53jpRW3TJCZ+vMXNeWzDz4DoNZyqecAAAAASUVORK5CYII=',
  LinkToFile: 0,
  SaveWithDocument: 0,
  Left: 30, // offset from left
  Top: 30, // offset from top
  Width: 32, // image width
  Height: 32 // image height
})
```

## Sort

Sort object; set the target range then apply sort.

#### Data type

[Sort](Sort.md) - Sort object

#### Example

```js
// Sort column C descending and column D ascending

// Current used range
const range = ActiveSheet.UsedRange
// Sort object
const sort = ActiveSheet.Sort
// Sort fields
const sortFields = sort.SortFields
// Clear previous sort keys
sortFields.Clear()
// Sort by column C values descending
sortFields.Add(Range('C:C').Item(1, 1), xlSortOnValues, xlDescending)
// Sort by column D values ascending
sortFields.Add(Range('D:D').Item(1, 1), xlSortOnValues, xlAscending)
// Header row: xlGuess auto, xlYes has header, xlNo no header
sort.Header = xlGuess
// Match case
sort.MatchCase = false
// Chinese sort: xlPinYin or xlStroke
sort.SortMethod = xlPinYin
// xlSortColumns vs xlSortRows
sort.Orientation = xlSortColumns
// SetRange required before Apply
sort.SetRange(range)
// Apply sort
sort.Apply()
```

## Range()

A cell or range as a Range (use Range properties and methods).

#### Data type

[Range](Range.md) - A cell or range on the sheet

#### Example

```js
const sheet = Application.ActiveSheet
// Print D2 text on active sheet
console.log(sheet.Range('D2').Text) // D2

// Set D2 value on active sheet
sheet.Range('D2').Value = 'this is D2'
```

## Cells()

Selects a cell on the sheet as a Range.

#### Return type

[Range](Range.md) - A cell on the sheet

#### Example

```js
// Print address of first cell on active sheet
console.log(Application.Cells(1).Address()) // $A$1
console.log(Application.Cells(2).Address()) // $A$2
```

## Activate()

Activates the worksheet

#### Example

```js
const sheet = Application.Sheets.Item(1)
// Activate first sheet; Application.ActiveSheet.Range and Application.Range refer to that sheet
sheet.Activate()
// Sets A1 on first sheet
Application.Range('A1').Value = 'foo'
```

## Move()

Moves the worksheet

#### Parameters

**The two parameters are mutually exclusive**

| Property | Data type | Default | Required | Description |
| --- | --- | --- | --- | --- |
| Before | number | null | No | Sheet id before which to place the moved sheet. Cannot use with After. |
| After | number | null | No | Sheet id after which to place the moved sheet. Cannot use with Before |

#### Example

```js
// Move current sheet after second sheet
const sheet = Application.ActiveSheet
sheet.Move({
  Before: null,
  After: Application.Sheets(2).Id
})
```

## Delete()

Deletes the worksheet

#### Return value

undefined

#### Example

```js
// Delete sheet named Sheet2
Application.Sheets.Item('Sheet2').Delete()
```
