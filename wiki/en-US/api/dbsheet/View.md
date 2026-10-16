---
title: "View"
source: https://airsheet.wps.cn/docs/en-US/api/dbsheet/View.html
updated: 1789700916000
hash: Dv45jbKc
---

# View

View operations.

## List views

```javascript
var views = Application.View.GetViews({ SheetId: 1 })
```

## Create a view

View types: [Multidimensional Table view types](/docs/en-US/api/dbsheet/Appendix.html#appendix-2-multidimensional-table-view-types).

```javascript
var view = Application.View.CreateView({
  SheetId: 1,
  Name: 'Grid',
  ViewType: 'Grid'
})
```

## Delete a view

```javascript
Application.View.DeleteView({ SheetId: 1, ViewId: 'F' })
```

## Rename a view

```javascript
Application.View.UpdateView({
  SheetId: 8,
  ViewId: 'E',
  Name: 'Roster'
})
```
