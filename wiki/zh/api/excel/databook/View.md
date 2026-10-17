---
title: "View"
source: https://airsheet.wps.cn/docs/api/excel/databook/View.html
updated: 1789700916000
hash: B38iwbZS
---

# View

视图操作

### 方法列表

| 方法名 | 返回类型 | 简介 |
| --- | --- | --- |
| [GetViews()](#getviews) | Array | 获取视图信息 |
| [CreateView()](#createview) | Object | 创建视图 |
| [DeleteView()](#deleteview) | Object | 删除视图 |
| [UpdateView()](#updateview) | Object | 修改视图名称 |

## GetViews()

获取视图信息

### 返回值

Array - 获取的表下所有视图的列表，单项视图信息中的字段说明如下所示：

| 属性 | 数据类型 | 说明 |
| --- | --- | --- |
| id | String | 视图Id |
| name | String | 视图名称 |
| recordsCount | Number | 视图的记录数量 |
| type | String | [视图类型](附录.md#附录-2-数据表视图类型说明) |

### 示例

```javascript
const sheet = Application.ActiveSheet
// 获取的表下所有视图的列表
const views = sheet.View.GetViews()
console.log(views)
// 打印结果：
// [
//     {"id":"Cm","name":"表","recordsCount":1,"type":"Grid"},
//     {"id":"Cn","name":"看板","recordsCount":1,"type":"Kanban"},
//     {"id":"Co","name":"画册","recordsCount":1,"type":"Gallery"},
// ]
```

## CreateView()

创建视图

### 参数

| 属性 | 数据类型 | 默认值 | 必填 | 说明 |
| --- | --- | --- | --- | --- |
| Name | String |  | 是 | 视图名称 |
| ViewType | String |  | 是 | [视图类型](附录.md#附录-2-数据表视图类型说明) |

### 返回值

Object - 返回内含新创建视图的各种信息

| 属性 | 数据类型 | 说明 |
| --- | --- | --- |
| id | String | 视图Id |
| name | String | 视图名称 |
| recordsCount | Number | 视图的记录数量 |
| type | String | [视图类型](附录.md#附录-2-数据表视图类型说明) |

### 示例

```javascript
const sheet = Application.ActiveSheet
// 获取的表下所有视图的列表
const view = sheet.View.CreateView({
    Name: '看板',
    ViewType: 'Kanban'
})
console.log(view)
// 打印结果：
// {"id":"Cu","name":"看板","recordsCount":18,"type":"Kanban"}
```

## DeleteView()

删除视图

### 参数

| 属性 | 数据类型 | 默认值 | 必填 | 说明 |
| --- | --- | --- | --- | --- |
| ViewId | String |  | 是 | 视图Id |

### 返回值

Object - 返回已删除的视图Id

| 属性 | 数据类型 | 说明 |
| --- | --- | --- |
| id | String | 视图Id |

### 示例

```javascript
const sheet = Application.ActiveSheet
// 获取的表下所有视图的列表
const views = sheet.View.GetViews()
// 获取待删除的视图id
const id = views[0].id
// 删除视图
const deleteSheet = sheet.View.DeleteView({ ViewId: id })
console.log(deleteSheet)
// 打印结果：
// {"id":"Cu"}
```

## UpdateView()

修改视图名称

### 参数

| 属性 | 数据类型 | 默认值 | 必填 | 说明 |
| --- | --- | --- | --- | --- |
| ViewId | String |  | 是 | 视图Id |
| Name | String |  | 是 | 更新的视图名称 |

### 返回值

Object - 返回已更新的视图信息

| 属性 | 数据类型 | 说明 |
| --- | --- | --- |
| id | String | 视图Id |
| name | String | 视图名称 |
| recordsCount | Number | 视图的记录数量 |
| type | String | [视图类型](附录.md#附录-2-数据表视图类型说明) |

### 示例

```javascript
const sheet = Application.ActiveSheet
// 获取的表下所有视图的列表
const views = sheet.View.GetViews()
// 获取待删除的视图id
const id = views[0].id
// 修改视图名称
const updateSheet = sheet.View.UpdateView({
    ViewId: id,
    Name: '名单'
})
console.log(updateSheet)
// 打印结果：
// {"id":"E","name":"名单","recordsCount":1,"type":"Grid"}
```
