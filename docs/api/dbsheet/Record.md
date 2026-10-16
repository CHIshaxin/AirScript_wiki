---
title: "行记录"
source: https://airsheet.wps.cn/docs/api/dbsheet/Record.html
---
# Record

行记录

## 获取行记录（多条）

可选参数列表

| 参数 | 说明 |
| --- | --- |
| ViewId | 填写后将从被指定的视图获取该用户所见到的记录；若不填写，则从工作表获取记录 |
| PageSize | 存在分页时，指定本次查询的起始记录（含）。若不填写或填写为空字符串，则从第一条记录开始获取 |
| Offset | 分页查询时，还将返回一个 `offset` 值，指向下一页的第一条记录，供后续查询。查询到最后一页或第 `maxRecords` 条记录时，返回数据将不再包含 `offset` 值 |
| MaxRecords | 指定要获取的“前 `maxRecords` 条记录”，若不填写，则默认返回全部记录 |
| Fields | ["数量", "日期"，"时间"]，指定所返回记录中的字段信息，若不填写，则默认返回全部字段内的信息。 |
| Filter | 筛选条件，详细说明见[筛选条件说明](../excel/databook/附录.md#附录-3-筛选条件说明) |

```javascript
var records = Application.Record.GetRecords({ SheetId: 3 })
```

## 获取行记录（单条）

```javascript
var record = Application.Record.GetRecord({ SheetId: 3, RecordId: 'Bz' })
```

## 删除行记录

```javascript
var result = Application.Record.DeleteRecords({
  SheetId: 3,
  RecordIds: ['J', 'P', 'Q']
})
```

## 更新行记录

Field 格式说明见 [多维表字段类型说明](附录.md#附录1：多维表字段类型说明)

```javascript
var records = recordapi.UpdateRecords({
  SheetId: 5,
  Records: [
    {
      id: 'A',
      fields: {
        邮箱: 'demo@qq.com',
        多选: ['1', '2']
      }
    }
  ]
})
```

## 创建行记录

Field 格式说明见 [多维表字段类型说明](附录.md#附录1：多维表字段类型说明)

```javascript
var records = recordapi.CreateRecords({
  SheetId: 5,
  Records: [
    {
      fields: {
        邮箱: 'demo@qq.com',
        多选: ['1', '2']
      }
    }
  ]
})
```
