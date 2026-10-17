---
title: "Record"
source: https://airsheet.wps.cn/docs/en-US/api/dbsheet/Record.html
updated: 1789700916000
hash: Bz_w8ISy
---

# Record

Row records.

## Get multiple row records

Optional parameters:

| Parameter | Description |
| --- | --- |
| ViewId | If set, records are read from that view as the user sees them; otherwise from the worksheet |
| PageSize | When paginating, the starting record (inclusive) for this query. If empty, starts from the first record |
| Offset | For pagination, the response may include `offset` for the first row of the next page. Omitted on the last page or when `maxRecords` is reached |
| MaxRecords | Return only the first `maxRecords` rows; if omitted, all matching rows |
| Fields | e.g. `["Quantity", "Date", "Time"]` – field names to return; if omitted, all fields |
| Filter | Filter criteria; see [filter criteria](Appendix.md#appendix-3-filter-criteria) |

```javascript
var records = Application.Record.GetRecords({ SheetId: 3 })
```

## Get a single row record

```javascript
var record = Application.Record.GetRecord({ SheetId: 3, RecordId: 'Bz' })
```

## Delete row records

```javascript
var result = Application.Record.DeleteRecords({
  SheetId: 3,
  RecordIds: ['J', 'P', 'Q']
})
```

## Update row records

Field formats: [Multidimensional Table field types](Appendix.md#appendix-1-multidimensional-table-field-types).

```javascript
var records = recordapi.UpdateRecords({
  SheetId: 5,
  Records: [
    {
      id: 'A',
      fields: {
        email: 'demo@gmail.com',
        multiSelect: ['1', '2']
      }
    }
  ]
})
```

## Create row records

Field formats: [Multidimensional Table field types](Appendix.md#appendix-1-multidimensional-table-field-types).

```javascript
var records = recordapi.CreateRecords({
  SheetId: 5,
  Records: [
    {
      fields: {
        email: 'demo@gmail.com',
        multiSelect: ['1', '2']
      }
    }
  ]
})
```
