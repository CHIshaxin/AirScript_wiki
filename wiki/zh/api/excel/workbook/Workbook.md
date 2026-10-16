---
title: "Workbook"
source: https://airsheet.wps.cn/docs/api/excel/workbook/Workbook.html
updated: 1789700916000
hash: DoaQPCQc
---

# Workbook

工作簿，文件的工作区,包含所有 Sheet

Workbook 对象的具体属性和方法请参阅下方的列表。

#### 属性列表

| 属性名 | 数据类型 | 简介 |
|---|---|---|
| [ActiveSheet](#activesheet) | [Sheet](/docs/api/excel/workbook/Sheet.html) | 工作簿中的活动工作表 |
| [Sheets](#sheets) | [Sheets](/docs/api/excel/workbook/Application.html#sheets) | 工作表的所有对象集合 |
| [ReadOnly](#readonly) | Boolean | 文档是否只读 |
| [ReadOnlyComment](#readonlycomment) | Boolean | 文档是否只读可评论的权限 |
| [SupportReadOnlyComment](#supportreadonlycomment) | Boolean | 文档是否支持只读可评论权限 |

#### 方法列表

| 方法名 | 返回类型 | 简介 |
|---|---|---|
| [Save()](#save) | String(JSON) | 保存文件的改动 |
| [GetComments()](#getcomments) | String(JSON) | 获取整个 Workbook 的评论 |
| [ExportAsFixedFormat()](#exportasfixedformat) | String(JSON) | 导出整个表格的 PDF 或者 Img 图片 |

## ActiveSheet

工作簿中的活动工作表

可用`Application.ActiveSheet`代替

#### 数据类型

[Sheet](/docs/api/excel/workbook/Sheet.html) - 活动工作簿中的活动工作表

#### 示例

```js
// 下面两种写法是一样的
console.log(Application.ActiveWorkbook.ActiveSheet.Name) // Sheet1
console.log(Application.ActiveSheet.Name) // Sheet1
```

## Sheets

获取当前文件能操作的所有Sheet，返回一个[Sheets](/docs/api/excel/workbook/Application.html#sheets)对象。

#### 数据类型

[Sheets](/docs/api/excel/workbook/Application.html#sheets)

#### 示例

```js
// 下面两种写法是一样的
let sheets = Application.ActiveWorkbook.Sheets
sheets = Application.Sheets
// 打印所有工作表的名称
for (let i = 1; i <= sheets.Count; i++) {
    console.log((sheets.Item(i).Name))
}
```

## ReadOnly

返回一个值，表示文档是否只读，此属性为只读属性。

#### 数据类型

Boolean - 文档是否只读

#### 示例

```js
// 打印当前文档是否只读
console.log(Application.ActiveWorkbook.ReadOnly) //false
```

## ReadOnlyComment

返回一个值，表示文档是否只读可评论的权限，此属性为只读属性。

#### 数据类型

Boolean - 文档是否只读可评论的权限

#### 示例

```js
// 打印当前文档是否只读可评论
console.log(Application.ActiveWorkbook.ReadOnlyComment) //true
```

## Save()

保存文件的改动

#### 返回类型

String(JSON) - 文件的保存状态，具体格式如下：

| 属性 | 数据类型 | 说明 |
|---|---|---|
| result | String | 保存状态 |
| size | Number | 文件大小，单位 byte |
| version | Number | 版本 |

**保存状态说明**：

| 保存状态 | 说明 |
|---|---|
| ok | 版本保存成功，可在历史版本中查看 |
| nochange | 文档无更新，无需保存版本 |
| SavedEmptyFile | 暂不支持保存空文件 触发场景：内核保存完后文件为空 |
| SpaceFull | 空间已满 |
| QueneFull | 保存中请勿频繁操作 触发场景：服务端处理保存队列已满，正在排队 |
| fail | 保存失败 |

#### 示例

```js
// 保存文件的改动
let res = Application.ActiveWorkbook.Save()
if (res.result === 'ok') {
  console.log('成功')
} else {
  console.error('失败:', res.result)
}
```

## GetComments()

获取整个 Workbook 的评论

#### 返回类型

String(JSON) - 整个 workbook 的评论集合，具体格式如下：

| 属性 | 数据类型 | 说明 |
|---|---|---|
| CellComments | `Array<Object>` | 评论信息集合 |
| PosInfo | `String` | 单元格信息 |
| SheetName | `String` | Sheet 名称 |
| UserIds | `Array<String>` | 用户 id 集合 |

**评论信息集合**：

| 属性 | 数据类型 | 说明 |
|---|---|---|
| DateTime | `String` | 时间戳 |
| Text | `String` | 评论文本 |
| Time | `String` | 转换后的时间 |
| UserId | `String` | 用户 id |

#### 示例

```js
// 获取整个 Workbook 的评论
let json = Application.ActiveWorkbook.GetComments()
for (let i = 0; i < json.length; i++) {
  let pos = json[i].PosInfo
  let comments = json[i].CellComments
  let commentList = []
  for (let j = 0; j < comments.length; j++) {
    commentList.push(comments[j].Text)
  }
  console.log(pos, ':', commentList)
}
```

## ExportAsFixedFormat()

导出整个表格为对应的 PDF 或者 Img 图片，并获取导出后的 url

#### 参数

| 属性 | 数据类型 | 默认值 | 必填 | 说明 |
|---|---|---|---|---|
| Type | Enum |  | 可选 | 导出的类型，详细可参考 [Enum.XlFixedFormatType](/docs/api/excel/workbook/Enum.html#xlfixedformattype)，目前仅支持导出图片和导出 PDF |

#### 返回类型

String(JSON) - 导出 URL 的 JSON 字符串

#### 示例 1：导出 PDF

```js
// 导出整个表格
let json = Application.ActiveWorkbook.ExportAsFixedFormat()
console.log(json.url)
```

#### 示例 2：导出图片

```js
// 导出整个表格
let json = Application.ActiveWorkbook.ExportAsFixedFormat({
  Type: Application.Enum.XlFixedFormatType.xlTypeIMG
})
console.log(json.url)
```
