---
title: "Field"
source: https://airsheet.wps.cn/docs/ja-JP/api/dbsheet/Field.html
updated: 1789700916000
hash: 62eOQ0X-
---

# Field

Field の操作です。

## フィールド定義の取得

```javascript
var fields = Application.Field.GetFields({ SheetId: 1 })
```

## フィールドの作成

フィールドの形式: [多次元テーブルのフィールド型](/docs/ja-JP/api/dbsheet/Appendix.html#appendix-1-multidimensional-table-field-types)。

```javascript
var field = Application.Field.CreateFields({
  SheetId: 3,
  Fields: [{ name: 'Score', type: 'Rating', max: 5 }]
})
```

## フィールドの削除

```javascript
// 戻り値には各削除の成否が含まれます
// [{"deleted":false,"id":"P"},{"deleted":false,"id":"Q"}]
var resutlt = Application.Field.DeleteFields({
  SheetId: 8,
  FieldIds: ['P', 'Q']
})
```

## フィールドの更新

```javascript
Application.Field.UpdateFields({
  SheetId: 3,
  Fields: [{ id: '6', name: 'Name' }]
})
```
