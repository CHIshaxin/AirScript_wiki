---
title: "Field"
source: https://airsheet.wps.cn/docs/en-US/api/excel/databook/Field.html
updated: 1789700916000
hash: c16k6sUJ
---

# Field

Field operations.

### Methods

| Method | Return type | Description |
| --- | --- | --- |
| [GetFields()](#getfields) | Array | Get field definitions |
| [CreateFields()](#createfields) | Array | Create fields |
| [DeleteFields()](#deletefields) | Array | Delete fields |
| [UpdateFields()](#updatefields) | Array | Update fields |

## GetFields()

Gets field definitions.

### Return value

Array – all fields on the sheet

| Property | Type | Description |
| --- | --- | --- |
| id | String | Field ID |
| name | String | Field name |
| type | String | [Field type](Appendix.md#appendix-1-datasheet-field-types) |

### Example

```javascript
const sheet = Application.ActiveSheet
const fields = sheet.Field.GetFields()
console.log(fields)
// [
//  {"id":"Ce","name":"Name","type":"MultiLineText"},
//  {"id":"Cf","name":"Quantity","type":"Number"},
// ]
```

## CreateFields()

Creates fields.

### Parameters

| Property | Type | Default | Required | Description |
| --- | --- | --- | --- | --- |
| Fields | Array |  | Yes | Field definitions; format see [Appendix](Appendix.md#appendix-1-datasheet-field-types) |

### Return value

Array – created fields

| Property | Type | Description |
| --- | --- | --- |
| id | String | Field ID |
| name | String | Field name |
| type | String | [Field type](Appendix.md#appendix-1-datasheet-field-types) |

### Example

```javascript
const sheet = Application.ActiveSheet
const field = sheet.Field.CreateFields({
  Fields: [{ name: 'Score', type: 'Rating', max: 5 }]
})
console.log(field)
// [{"id":"LZ","name":"Score","type":"Rating"}]
```

## DeleteFields()

Deletes fields.

### Parameters

| Property | Type | Default | Required | Description |
| --- | --- | --- | --- | --- |
| Fields | Array |  | Yes | Field IDs to delete |

### Return value

Array – per-field delete result

| Property | Type | Description |
| --- | --- | --- |
| id | String | Field ID |
| deleted | Boolean | Whether deletion succeeded |

### Example

```javascript
const sheet = Application.ActiveSheet
const resutlt = sheet.Field.DeleteFields({ FieldIds: ['P', 'Q'] })
console.log(resutlt)
// [{"deleted":false,"id":"P"},{"deleted":false,"id":"Q"}]
```

## UpdateFields()

Updates fields.

### Parameters

| Property | Type | Default | Required | Description |
| --- | --- | --- | --- | --- |
| Fields | Array |  | Yes | Updates including field ID and name; format see [Appendix](Appendix.md#appendix-1-datasheet-field-types) |

### Return value

Array – updated field definitions

| Property | Type | Description |
| --- | --- | --- |
| id | String | Field ID |
| name | String | Field name |
| type | String | [Field type](Appendix.md#appendix-1-datasheet-field-types) |

### Example

```javascript
const sheet = Application.ActiveSheet
sheet.Field.UpdateFields({
    Fields: [{ id: 'LG', name: 'Jump' }]
})
```
