---
title: "Record"
source: https://airsheet.wps.cn/docs/api/excel/databook/Record.html
updated: 1789700916000
hash: DbbC9VTX
---

# Record

行记录

### 方法列表

| 方法名 | 返回类型 | 简介 |
| --- | --- | --- |
| [GetRecords()](#getrecords) | Array | 获取行记录（多条） |
| [GetRecord()](#getrecord) | Object | 获取行记录（单条） |
| [DeleteRecords()](#deleterecords) | Array | 删除行记录 |
| [UpdateRecords()](#updaterecords) | Array | 更新行记录 |
| [CreateRecords()](#createrecords) | Array | 创建行记录 |
| [GetAttachmentURL()](#getattachmenturl) | String | 获取上传附件或图片的URL |

## GetRecords()

获取行记录（多条）

注意

每次请求最多返回100条，数据量大的时候请使用分页查询

### 参数

| 属性 | 数据类型 | 默认值 | 必填 | 说明 |
| --- | --- | --- | --- | --- |
| ViewId | String |  | 否 | 填写后将从被指定的视图获取该用户所见到的记录；若不填写，则从工作表获取记录 |
| PageSize | Number | 100 | 否 | 存在分页时，指定本次查询的起始记录（含）。若不填写或填写为空字符串，则从第一条记录开始获取。**当前最大值：1000** |
| Offset | Number |  | 否 | 分页查询时，将返回一个offset值，指向下一页的第一条记录，供后续查询。查询到最后一页或第maxRecords条记录时，返回数据将不再包含offset值 |
| MaxRecords | Number |  | 否 | 指定要获取的“前maxRecords条记录”，若不填写，则默认返回全部记录 |
| Fields | Array |  | 否 | [字段类型](附录.md#附录-1-数据表字段类型说明) |
| Filter | Object |  | 否 | 详细说明见[附录三](附录.md#附录-3-筛选条件说明) |

### 返回值

Object - 获取表的所有记录

| 属性 | 数据类型 | 说明 |
| --- | --- | --- |
| Offset | String | 如果分页的话， 则会返回此字段信息;分页截止 id， 下次请求携带会继续分页请求信息 |
| Records | Array[Object] | 记录集合 |

#### 记录集合

| 属性 | 数据类型 | 说明 |
| --- | --- | --- |
| id | String | 记录Id |
| Fields | Object | 更新的字段信息，包含字段Id，字段name,格式说明见[附录](附录.md#附录-1-数据表字段类型说明) |

### 示例

```javascript
const sheet = Application.ActiveSheet
// 分页查询例子
function fetchAllRecords() {
  const view = sheet.Selection.GetActiveView()
  let all = []
  let offset = null;

  while (all.length === 0 || offset) {
    let records = sheet.Record.GetRecords({
      ViewId: view.viewId,
      Offset: offset,
    })
    offset = records.offset
    all = all.concat(records.records)
  }
  console.log(all.length)
  return all
}

fetchAllRecords()
```

## GetRecord()

获取行记录（单条）

### 参数

| 属性 | 数据类型 | 默认值 | 必填 | 说明 |
| --- | --- | --- | --- | --- |
| RecordId | String |  | 是 | 表中指定获取的记录id |

### 返回值

Object - 获取表的指定的单条记录

| 属性 | 数据类型 | 说明 |
| --- | --- | --- |
| id | String | 记录Id |
| Fields | Object | 更新的字段信息，包含字段Id，字段name,格式说明见[附录](附录.md#附录-1-数据表字段类型说明) |

### 示例

```javascript
const sheet = Application.ActiveSheet
const record = sheet.Record.GetRecord({  RecordId: 'Bz' })
console.log(record)
// 打印结果：
//  {"fields":{"日期":"2023/02/21"},"id":"Bz"}
```

## DeleteRecords()

删除行记录

### 参数

| 属性 | 数据类型 | 默认值 | 必填 | 说明 |
| --- | --- | --- | --- | --- |
| RecordIds | Array |  | 是 | 表中需要删除的记录id |

### 返回值

Array - 返回删除的表id以及删除是否成功信息

| 属性 | 数据类型 | 说明 |
| --- | --- | --- |
| id | String | 记录Id |
| deleted | Boolean | 是否删除成功 “true”表示删除成功，“false”表示删除失败 |

### 示例

```javascript
const sheet = Application.ActiveSheet
const result = sheet.Record.DeleteRecords({
    RecordIds: ['J', 'P', 'Q']
})
console.log(resutlt)
// 打印结果：
// [{"deleted":true,"id":"P"},{"deleted":false,"id":"Q"}]
```

## UpdateRecords()

更新行记录

### 参数

| 属性 | 数据类型 | 默认值 | 必填 | 说明 |
| --- | --- | --- | --- | --- |
| Records | Array[Object] |  | 是 | 行记录集合 |

#### 行记录集合：

| 属性 | 数据类型 | 说明 |
| --- | --- | --- |
| id | String | 记录Id |
| Fields | Object | 更新的字段信息，包含字段Id，字段name,格式说明见[附录](附录.md#附录-1-数据表字段类型说明) |

### 返回值

Array - 表的已更新的所有记录

| 属性 | 数据类型 | 说明 |
| --- | --- | --- |
| id | String | 记录Id |
| Fields | Object | 更新的字段信息，包含字段Id，字段name,格式说明见[附录](附录.md#附录-1-数据表字段类型说明) |

### 示例

```javascript
const sheet = Application.ActiveSheet
const records = sheet.Record.UpdateRecords({
        Records: [{
            id: 'A',
            fields: {
                 邮箱: 'demo@qq.com',
                 多选: ['1', '2'],
                 "记录关联": {
                    "recordIds": ["I", "K"]
                 }
            }
        }],
    })
```

## CreateRecords()

创建行记录

### 参数

| 属性 | 数据类型 | 默认值 | 必填 | 说明 |
| --- | --- | --- | --- | --- |
| Records | Array[Object] |  | 是 | 行记录集合 |

#### 行记录集合：

| 属性 | 数据类型 | 说明 |
| --- | --- | --- |
| id | String | 记录Id |
| Fields | Object | 更新的字段信息，包含字段Id，字段name,格式说明见[附录](附录.md#附录-1-数据表字段类型说明) |

### 返回值

Array - 表的已更新的所有记录

| 属性 | 数据类型 | 说明 |
| --- | --- | --- |
| id | String | 记录Id |
| Fields | Object | 更新的字段信息，包含字段Id，字段name,格式说明见[附录](附录.md#附录-1-数据表字段类型说明) |

### 示例

```javascript
const sheet = Application.ActiveSheet
// 创建邮箱和多选
const records = sheet.Record.CreateRecords({
      Records: [{
          fields: {
               邮箱: 'demo@qq.com',
               多选: ['1', '2'],
          }
      }, {
          fields: {
               邮箱: 'demo@qq.com',
               多选: ['1', '2'],
          }
      }],
  })

// 创建联系人
const records = sheet.Record.CreateRecords({
  Records: [
    {  fields: { '联系人': [{ name: 'yourname', nickName: 'yourname', id: '88888888', avatar_url: 'https://avatar.qwps.cn/avatar/5b2t57-U' }] } },
  ],
});
```

## GetAttachmentURL()

获取上传附件或图片的URL

### 参数

注意

必须至少传入1个参数`Attachment` 或者传入2个参数`UploadId`和 `Source`

| 属性 | 数据类型 | 默认值 | 必填 | 说明 |
| --- | --- | --- | --- | --- |
| Attachment | String |  | 否 | 附件 |
| UploadId | String |  | 否 | 上传文件id |
| Source | String |  | 否 | source参数必须为"upload_ks3"（本地上传）或"cloud"（云上传） |

### 返回值

String - 为获取上传附件或图片的URL，打开该URL可进行附件或图片下载

### 示例

```javascript
const sheet = Application.ActiveSheet
const resultURL = sheet.Record.GetAttachmentURL({
    Attachment: "IKWRCBAAKA|upload_ks3|image/png|QQ图片20230214165215.png|12070||549*106",
      })

//or

const resultURL = sheet.Record.GetAttachmentURL({
    UploadId: "IKWRCBAAKA",
    Source: "upload_ks3"
      })
```
