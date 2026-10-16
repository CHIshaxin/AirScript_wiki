---
title: "Sheet"
source: https://airsheet.wps.cn/docs/en-US/api/dbsheet/Sheet.html
updated: 1789700916000
hash: BGSvpjHc
---

# Sheet

Sheet operations.

## List all sheets

```javascript
var sheets = Application.Sheet.GetSheets()
```

## Create a sheet

Parameters

- Pass at least one sheet and one field.
- Field formats: [Multidimensional Table field types](/docs/en-US/api/dbsheet/Appendix.html#appendix-1-multidimensional-table-field-types).
- View types: [Multidimensional Table view types](/docs/en-US/api/dbsheet/Appendix.html#appendix-2-multidimensional-table-view-types).

```javascript
var sheet = Application.Sheet.CreateSheet({
  Name: 'Products',
  Views: [{ name: 'Grid', type: 'Grid' }],
  Fields: [
    { name: 'Name', type: 'MultiLineText' },
    {
      name: 'Category',
      type: 'SingleSelect',
      items: [{ value: 'A' }, { value: 'B' }, { value: 'C' }]
    }
  ]
})
```

## Delete a sheet

```javascript
Application.Sheet.DeleteSheet({ SheetId: 8 })
```

## Rename a sheet

```javascript
var result = Application.Sheet.UpdateSheet({ SheetId: 9, Name: 'Backup' })
```
