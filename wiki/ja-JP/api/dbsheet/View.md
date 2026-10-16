---
title: "View"
source: https://airsheet.wps.cn/docs/ja-JP/api/dbsheet/View.html
updated: 1789700916000
hash: Zwe72BZw
---

# View

View の操作です。

## View の一覧

```javascript
var views = Application.View.GetViews({ SheetId: 1 })
```

## View の作成

View の型: [多次元テーブルの View 型](/docs/ja-JP/api/dbsheet/Appendix.html#appendix-2-multidimensional-table-view-types)。

```javascript
var view = Application.View.CreateView({
  SheetId: 1,
  Name: 'Grid',
  ViewType: 'Grid'
})
```

## View の削除

```javascript
Application.View.DeleteView({ SheetId: 1, ViewId: 'F' })
```

## View の名前変更

```javascript
Application.View.UpdateView({
  SheetId: 8,
  ViewId: 'E',
  Name: 'Roster'
})
```
