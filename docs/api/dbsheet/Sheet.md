---
title: "表"
source: https://airsheet.wps.cn/docs/api/dbsheet/Sheet.html
---
# Sheet

表操作

## 获取所有表信息

```javascript
var sheets = Application.Sheet.GetSheets()
```

## 创建新的表

> **参数说明**
>
> - 最少传入一个表和一个字段
> - Field 格式说明见 [多维表字段类型说明](附录.md#附录1：多维表字段类型说明)
> - 视图类型说明见 [多维表视图类型说明](附录.md#附录2：多维表视图类型说明#多维表视图类型说明)

```javascript
var sheet = Application.Sheet.CreateSheet({
  Name: '商品表',
  Views: [{ name: '表格', type: 'Grid' }],
  Fields: [
    { name: '名称', type: 'MultiLineText' },
    {
      name: '分类',
      type: 'SingleSelect',
      items: [{ value: 'A' }, { value: 'B' }, { value: 'C' }]
    }
  ]
})
```

## 删除表

```javascript
Application.Sheet.DeleteSheet({ SheetId: 8 })
```

## 更新表名字

```javascript
var result = Application.Sheet.UpdateSheet({ SheetId: 9, Name: '备份表' })
```
