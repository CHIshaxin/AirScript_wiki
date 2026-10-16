---
title: "Field"
source: https://airsheet.wps.cn/docs/en-US/api/dbsheet/Field.html
updated: 1789700916000
hash: BX2vAaN9
---

# Field

Field operations.

## Get field definitions

```javascript
var fields = Application.Field.GetFields({ SheetId: 1 })
```

## Create fields

Field formats: [Multidimensional Table field types](/docs/en-US/api/dbsheet/Appendix.html#appendix-1-multidimensional-table-field-types).

```javascript
var field = Application.Field.CreateFields({
  SheetId: 3,
  Fields: [{ name: 'Score', type: 'Rating', max: 5 }]
})
```

## Delete fields

```javascript
// result includes whether each delete succeeded
// [{"deleted":false,"id":"P"},{"deleted":false,"id":"Q"}]
var resutlt = Application.Field.DeleteFields({
  SheetId: 8,
  FieldIds: ['P', 'Q']
})
```

## Update fields

```javascript
Application.Field.UpdateFields({
  SheetId: 3,
  Fields: [{ id: '6', name: 'Name' }]
})
```
