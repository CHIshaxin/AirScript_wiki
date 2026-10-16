---
title: "View"
source: https://airsheet.wps.cn/docs/api/dbsheet/View.html
updated: 1789700916000
hash: DBgHdno0
---

# View

视图操作

## 获取视图信息

```javascript
var views = Application.View.GetViews({ SheetId: 1 }）
```

## 创建视图

视图类型说明见 [多维表视图类型说明](/docs/api/dbsheet/附录.html#附录2：多维表视图类型说明#多维表视图类型说明)

```javascript
var view = Application.View.CreateView({
  SheetId: 1,
  Name: '表格',
  ViewType: 'Grid'
})
```

## 删除视图

```javascript
Application.View.DeleteView({ SheetId: 1, ViewId: 'F' })
```

## 修改视图名称

```javascript
Application.View.UpdateView({
  SheetId: 8,
  ViewId: 'E',
  Name: '名单'
})
```
