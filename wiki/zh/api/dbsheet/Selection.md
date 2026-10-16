---
title: "Selection"
source: https://airsheet.wps.cn/docs/api/dbsheet/Selection.html
updated: 1789700916000
hash: kQYl0FhD
---

# Selection

选区

## 获取当前选中视图

```javascript
var view = selectionHandler.GetActiveView()
```

## 获取选中记录信息

a. 这里返回的 records 是一个二维数组，因为选中记录可以是不连续的，如果是连续的记录，那么可以直接去 records[0]

```javascript
var records = selectionHandler.GetSelectionRecords()
```

## 设置选中的视图

```javascript
selectionHandler.SetActiveView({ SheetId: 27, ViewId: 'Q' })
```

## 获取当前选中表

```javascript
var sheet = selectionHandler.GetActiveSheet()
```
