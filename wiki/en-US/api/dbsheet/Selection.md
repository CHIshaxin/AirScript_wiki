---
title: "Selection"
source: https://airsheet.wps.cn/docs/en-US/api/dbsheet/Selection.html
updated: 1789700916000
hash: CcHAU_tN
---

# Selection

Selection range.

## Get the active view

```javascript
var view = selectionHandler.GetActiveView()
```

## Get selected records

The returned `records` is a two-dimensional array because the selection can be non-contiguous. If the selection is contiguous, you can use `records[0]` directly.

```javascript
var records = selectionHandler.GetSelectionRecords()
```

## Set the active view

```javascript
selectionHandler.SetActiveView({ SheetId: 27, ViewId: 'Q' })
```

## Get the active sheet

```javascript
var sheet = selectionHandler.GetActiveSheet()
```
