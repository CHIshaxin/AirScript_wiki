---
title: "Record"
source: https://airsheet.wps.cn/docs/en-US/api/excel/databook/Record.html
updated: 1789700916000
hash: N0Zn-ZP1
---

# Record

Row records.

### Methods

| Method | Return type | Description |
|---|---|---|
| [GetRecords()](#getrecords) | Array | Get multiple rows |
| [GetRecord()](#getrecord) | Object | Get a single row |
| [DeleteRecords()](#deleterecords) | Array | Delete rows |
| [UpdateRecords()](#updaterecords) | Array | Update rows |
| [CreateRecords()](#createrecords) | Array | Create rows |
| [GetAttachmentURL()](#getattachmenturl) | String | URL for uploaded attachment or image |

## GetRecords()

Gets multiple row records.

Note

Each request returns at most 100 rows. Use pagination for large datasets.

### Parameters

| Property | Type | Default | Required | Description |
|---|---|---|---|---|
| ViewId | String |  | No | If set, records are read from that view as the user sees them; otherwise from the sheet |
| PageSize | Number | 100 | No | When paginating, the starting record (inclusive) for this query. If empty, starts from the first record. **Current maximum: 1000** |
| Offset | Number |  | No | For pagination, the response may include `offset` for the first row of the next page. Omitted on the last page or when `maxRecords` is reached |
| MaxRecords | Number |  | No | Return only the first `maxRecords` rows; if omitted, all matching rows |
| Fields | Array |  | No | [Field types](/docs/en-US/api/excel/databook/Appendix.html#appendix-1-datasheet-field-types) |
| Filter | Object |  | No | See [Appendix 3](/docs/en-US/api/excel/databook/Appendix.html#appendix-3-filter-criteria) |

### Return value

Object – batch of records

| Property | Type | Description |
|---|---|---|
| Offset | String | Pagination cursor; pass on the next request to continue |
| Records | Array[Object] | Records |

#### Record shape

| Property | Type | Description |
|---|---|---|
| id | String | Record ID |
| Fields | Object | Field values by ID or name; format see [Appendix](/docs/en-US/api/excel/databook/Appendix.html#appendix-1-datasheet-field-types) |

### Example

```javascript
const sheet = Application.ActiveSheet
function fetchAllRecords() {
  const view = sheet.Selection.GetActiveView()
  let all = []
  let offset = null

  while (all.length === 0 || offset) {
    let records = sheet.Record.GetRecords({
      ViewId: view.viewId,
      Offset: offset
    })
    offset = records.offset
    all = all.concat(records.records)
  }
  console.log(all.length)
  return all
}

fetchAllRecords()
```

## GetRecord()

Gets a single row.

### Parameters

| Property | Type | Default | Required | Description |
|---|---|---|---|---|
| RecordId | String |  | Yes | Record ID |

### Return value

Object – one record

| Property | Type | Description |
|---|---|---|
| id | String | Record ID |
| Fields | Object | Field values; format see [Appendix](/docs/en-US/api/excel/databook/Appendix.html#appendix-1-datasheet-field-types) |

### Example

```javascript
const sheet = Application.ActiveSheet
const record = sheet.Record.GetRecord({ RecordId: 'Bz' })
console.log(record)
// {"fields":{"Date":"2023/02/21"},"id":"Bz"}
```

## DeleteRecords()

Deletes rows.

### Parameters

| Property | Type | Default | Required | Description |
|---|---|---|---|---|
| RecordIds | Array |  | Yes | IDs to delete |

### Return value

Array – per-ID result

| Property | Type | Description |
|---|---|---|
| id | String | Record ID |
| deleted | Boolean | `true` if deleted, `false` if not |

### Example

```javascript
const sheet = Application.ActiveSheet
const result = sheet.Record.DeleteRecords({
  RecordIds: ['J', 'P', 'Q']
})
console.log(resutlt)
// [{"deleted":true,"id":"P"},{"deleted":false,"id":"Q"}]
```

## UpdateRecords()

Updates rows.

### Parameters

| Property | Type | Default | Required | Description |
|---|---|---|---|---|
| Records | Array[Object] |  | Yes | Records to update |

#### Record entry

| Property | Type | Description |
|---|---|---|
| id | String | Record ID |
| Fields | Object | Values by field ID or name; format see [Appendix](/docs/en-US/api/excel/databook/Appendix.html#appendix-1-datasheet-field-types) |

### Return value

Array – updated records

| Property | Type | Description |
|---|---|---|
| id | String | Record ID |
| Fields | Object | Updated values; format see [Appendix](/docs/en-US/api/excel/databook/Appendix.html#appendix-1-datasheet-field-types) |

### Example

```javascript
const sheet = Application.ActiveSheet
const records = sheet.Record.UpdateRecords({
  Records: [
    {
      id: 'A',
      fields: {
        email: 'demo@qq.com',
        multiSelect: ['1', '2'],
        recordLink: {
          recordIds: ['I', 'K']
        }
      }
    }
  ]
})
```

## CreateRecords()

Creates rows.

### Parameters

| Property | Type | Default | Required | Description |
|---|---|---|---|---|
| Records | Array[Object] |  | Yes | New records |

#### Record entry

| Property | Type | Description |
|---|---|---|
| id | String | Record ID (when applicable) |
| Fields | Object | Values; format see [Appendix](/docs/en-US/api/excel/databook/Appendix.html#appendix-1-datasheet-field-types) |

### Return value

Array – created records

| Property | Type | Description |
|---|---|---|
| id | String | Record ID |
| Fields | Object | Values; format see [Appendix](/docs/en-US/api/excel/databook/Appendix.html#appendix-1-datasheet-field-types) |

### Example

```javascript
const sheet = Application.ActiveSheet
const records = sheet.Record.CreateRecords({
  Records: [
    {
      fields: {
        email: 'demo@qq.com',
        multiSelect: ['1', '2']
      }
    },
    {
      fields: {
        email: 'demo@qq.com',
        multiSelect: ['1', '2']
      }
    }
  ]
})

const records2 = sheet.Record.CreateRecords({
  Records: [
    {
      fields: {
        contact: [
          {
            name: 'yourname',
            nickName: 'yourname',
            id: '88888888',
            avatar_url: 'https://example.com/avatar/user123'
          }
        ]
      }
    }
  ]
})
```

## GetAttachmentURL()

Gets a download URL for an attachment or image.

### Parameters

Note

Pass at least `Attachment`, or both `UploadId` and `Source`.

| Property | Type | Default | Required | Description |
|---|---|---|---|---|
| Attachment | String |  | No | Attachment token string |
| UploadId | String |  | No | Upload file ID |
| Source | String |  | No | Must be `"upload_ks3"` (local upload) or `"cloud"` (cloud upload) |

### Return value

String – URL to open for download

### Example

```javascript
const sheet = Application.ActiveSheet
const resultURL = sheet.Record.GetAttachmentURL({
  Attachment:
    'IKWRCBAAKA|upload_ks3|image/png|screenshot_20230214165215.png|12070||549*106'
})

const resultURL2 = sheet.Record.GetAttachmentURL({
  UploadId: 'IKWRCBAAKA',
  Source: 'upload_ks3'
})
```
