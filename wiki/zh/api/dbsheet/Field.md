---
title: "Field"
source: https://airsheet.wps.cn/docs/api/dbsheet/Field.html
updated: 1789700916000
hash: VJfuQyyw
---

# Field

字段操作

## 获取字段信息

```javascript
var fields = Application.Field.GetFields({ SheetId: 1 })
```

## 创建字段

Field 格式说明见 [多维表字段类型说明](/docs/api/dbsheet/附录.html#附录1：多维表字段类型说明)

```javascript
var field = Application.Field.CreateFields({
  SheetId: 3,
  Fields: [{ name: '等级', type: 'Rating', max: 5 }]
})
```

## 删除字段

```javascript
// result 中会返回是否删除成功
// [{"deleted":false,"id":"P"},{"deleted":false,"id":"Q"}]
var resutlt = Application.Field.DeleteFields({
  SheetId: 8,
  FieldIds: ['P', 'Q']
})
```

## 更新字段

```javascript
Application.Field.UpdateFields({
  SheetId: 3,
  Fields: [{ id: '6', name: '名称' }]
})
```
