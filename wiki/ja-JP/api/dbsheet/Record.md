---
title: "Record"
source: https://airsheet.wps.cn/docs/ja-JP/api/dbsheet/Record.html
updated: 1789700916000
hash: CAQx6DRv
---

# Record

行のレコードです。

## 複数行レコードの取得

任意のパラメーター:

| パラメーター | 説明 |
|---|---|
| ViewId | 指定すると、その View 上でユーザーが見ているとおりにレコードを読み取ります。未指定の場合はワークシートから読み取ります |
| PageSize | ページング時、このクエリで取得を開始するレコード（先頭を含む）。空の場合は先頭のレコードから開始します |
| Offset | ページング用。レスポンスに次ページ先頭行の `offset` が含まれることがあります。最終ページ、または `maxRecords` に達した場合は省略されます |
| MaxRecords | 一致する行の先頭から最大 `maxRecords` 行のみ返します。省略時は一致する行をすべて返します |
| Fields | 例: `["Quantity", "Date", "Time"]` — 返すフィールド名。省略時はすべてのフィールド |
| Filter | フィルター条件。[フィルター条件](/docs/ja-JP/api/dbsheet/Appendix.html#appendix-3-filter-criteria) を参照 |

```javascript
var records = Application.Record.GetRecords({ SheetId: 3 })
```

## 単一行レコードの取得

```javascript
var record = Application.Record.GetRecord({ SheetId: 3, RecordId: 'Bz' })
```

## 行レコードの削除

```javascript
var result = Application.Record.DeleteRecords({
  SheetId: 3,
  RecordIds: ['J', 'P', 'Q']
})
```

## 行レコードの更新

フィールドの形式: [多次元テーブルのフィールド型](/docs/ja-JP/api/dbsheet/Appendix.html#appendix-1-multidimensional-table-field-types)。

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

## 行レコードの作成

フィールドの形式: [多次元テーブルのフィールド型](/docs/ja-JP/api/dbsheet/Appendix.html#appendix-1-multidimensional-table-field-types)。

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
