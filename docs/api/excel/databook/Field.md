---
title: "字段(Field)"
source: https://airsheet.wps.cn/docs/api/excel/databook/Field.html
---
# Field

字段操作

### 方法列表

| 方法名 | 返回类型 | 简介 |
| --- | --- | --- |
| GetFields() | Array | 获取字段信息 |
| CreateFields() | Array | 创建字段 |
| DeleteFields() | Array | 删除字段 |
| UpdateFields() | Array | 更新字段 |

## GetFields()

获取字段信息

### 返回值

Array - 返回获取的表所有字段信息

| 属性 | 数据类型 | 说明 |
| --- | --- | --- |
| id | String | 字段Id |
| name | String | 字段名称 |
| type | String | [字段类型](附录.md#附录-1-数据表字段类型说明) |

### 示例

```javascript
const sheet = Application.ActiveSheet
// 获取的表所有字段信息
const fields = sheet.Field.GetFields()
console.log(fields)
// 打印结果：
// [
//  {"id":"Ce","name":"名称","type":"MultiLineText"},
//  {"id":"Cf","name":"数量","type":"Number"},
// ]
```

## CreateFields()

创建字段

### 参数

| 属性 | 数据类型 | 默认值 | 必填 | 说明 |
| --- | --- | --- | --- | --- |
| Fields | Array |  | 是 | 表的字段信息,格式说明见[附录](附录.md#附录-1-数据表字段类型说明) |

### 返回值

Array - 返回已创建的表所有字段信息

| 属性 | 数据类型 | 说明 |
| --- | --- | --- |
| id | String | 字段Id |
| name | String | 字段名称 |
| type | String | [字段类型](附录.md#附录-1-数据表字段类型说明) |

### 示例

```javascript
const sheet = Application.ActiveSheet
const field =  sheet.Field.CreateFields({
    Fields: [
        { name: '等级',  type: 'Rating', max: 5 }
    ]
})
console.log(field)
// 打印结果：
// [{"id":"LZ","name":"等级","type":"Rating"}]
```

## DeleteFields()

删除字段

### 参数

| 属性 | 数据类型 | 默认值 | 必填 | 说明 |
| --- | --- | --- | --- | --- |
| Fields | Array |  | 是 | 需要删除的字段Id |

### 返回值

Array - 返回删除的表id以及删除是否成功信息

| 属性 | 数据类型 | 说明 |
| --- | --- | --- |
| id | String | 字段Id |
| deleted | Boolean | 是否删除成功 |

### 示例

```javascript
const sheet = Application.ActiveSheet
// 删除字段
const resutlt = sheet.Field.DeleteFields({ FieldIds: ['P', 'Q'] })
console.log(resutlt)
// 打印结果：
// [{"deleted":false,"id":"P"},{"deleted":false,"id":"Q"}]
```

## UpdateFields()

更新字段

### 参数

| 属性 | 数据类型 | 默认值 | 必填 | 说明 |
| --- | --- | --- | --- | --- |
| Fields | Array |  | 是 | 更新的字段信息，包含字段Id，字段name,格式说明见[附录](附录.md#附录-1-数据表字段类型说明) |

### 返回值

Array - 返回已更新的字段信息

| 属性 | 数据类型 | 说明 |
| --- | --- | --- |
| id | String | 字段Id |
| name | String | 字段名称 |
| type | String | [字段类型](附录.md#附录-1-数据表字段类型说明) |

### 示例

```javascript
const sheet = Application.ActiveSheet
// 修改字段名称
sheet.Field.UpdateFields({
    Fields: [{ id: 'LG', name: '跳转' }]
})
```
