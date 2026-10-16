---
title: "Selection"
source: https://airsheet.wps.cn/docs/ja-JP/api/dbsheet/Selection.html
updated: 1789700916000
hash: CseelVyT
---

# Selection

選択範囲です。

## アクティブな View の取得

```javascript
var view = selectionHandler.GetActiveView()
```

## 選択中のレコードの取得

戻り値の `records` は、選択が連続していない場合があるため 2 次元配列です。選択が連続している場合は、`records[0]` をそのまま使えます。

```javascript
var records = selectionHandler.GetSelectionRecords()
```

## アクティブな View の設定

```javascript
selectionHandler.SetActiveView({ SheetId: 27, ViewId: 'Q' })
```

## アクティブな Sheet の取得

```javascript
var sheet = selectionHandler.GetActiveSheet()
```
