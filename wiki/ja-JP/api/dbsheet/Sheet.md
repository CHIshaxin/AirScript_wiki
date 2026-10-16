---
title: "Sheet"
source: https://airsheet.wps.cn/docs/ja-JP/api/dbsheet/Sheet.html
updated: 1789700916000
hash: C78voC7v
---

# Sheet

Sheet の操作です。

## すべての Sheet の一覧

```javascript
var sheets = Application.Sheet.GetSheets()
```

## Sheet の作成

パラメーター

- 少なくとも 1 つの Sheet と 1 つの Field を渡してください。
- フィールドの形式: [多次元テーブルのフィールド型](/docs/ja-JP/api/dbsheet/Appendix.html#appendix-1-multidimensional-table-field-types)。
- View の型: [多次元テーブルの View 型](/docs/ja-JP/api/dbsheet/Appendix.html#appendix-2-multidimensional-table-view-types)。

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

## Sheet の削除

```javascript
Application.Sheet.DeleteSheet({ SheetId: 8 })
```

## Sheet の名前変更

```javascript
var result = Application.Sheet.UpdateSheet({ SheetId: 9, Name: 'Backup' })
```
