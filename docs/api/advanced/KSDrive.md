---
title: "云文档 API"
source: https://airsheet.wps.cn/docs/api/advanced/KSDrive.html
---
# 云文档 API

AirScript 提供全局的 KSDrive 对象，通过此对象即可轻松**查看、修改和创建**您的云文档

> **提示**
>
> 在使用 KSDrive 对象操作云文档时，确保您已添加`云文档API`服务，在脚本编辑器的服务菜单内添加即可。

### 快速使用

```js
// 打开指定文档
let file = KSDrive.openFile('https://www.kdocs.cn/l/xxxxxxxxxxxx')
// 打印指定文档的A1单元格内容
console.log(file.Application.Range('A1').Text)
// 使用结束之后调用close关闭文档，否则无法再次调用KSDrive.openFile
file.close()
// 获取我的云文档下面的et，ksheet文档列表
const fileList = KSDrive.listFiles({ includeExts: ['et', 'ksheet'] })
// 打开我的云文档目录下的第一个文档
file = KSDrive.openFile(fileList.files[0])
console.log(file.Application.Range('A1').Text)
// 关闭文档
file.close()
```

### 属性列表

| 属性名 | 数据类型 | 说明 |
| --- | --- | --- |
| FileType | object | 支持的文件类型集合 |

### 方法列表

| 方法名 | 返回类型 | 说明 |
| --- | --- | --- |
| createFile() | string | 创建或另存一个文件 |
| openFile() | File | 额外打开一个文件 |
| listFiles() | FilesInfo | 列出某个目录下的表格文件 |

## FileType

云文档支持的文件类型，可用于新建文件时指定新文件的类型

### 属性说明

| 属性名 | 数据类型 | 说明 |
| --- | --- | --- |
| AP | string | 智能文档 |
| KSheet | string | 智能表格 |
| ET | string | 表格 |
| DB | string | 多维表 |

## createFile()

创建一个新文件，也可以将一个源文件另存为新文件

### 参数

| 名称 | 类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| type | FileType | 是 | 新文件的类型 |
| createOptions | CreateOptions | 是 | 新文件的参数选项 |

### CreateOptions 对象说明

| 名称 | 类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| name | string | 是 | 新文件的文件名 |
| dirUrl | string | 否 | 新文件的文件目录 |
| source | string | 否 | 将目标文件另存为新文件 |

### 返回值

url - string 新文件的 URL

### 示例

```js
// 创建ET文件，指定保存位置
let url = KSDrive.createFile(KSDrive.FileType.ET, {
  name: 'et测试',
  dirUrl: '指定保存位置'
})
console.log(url)
// 新建DB文件
url = KSDrive.createFile(KSDrive.FileType.DB)
console.log(url)
// 新建KSheet文件
url = KSDrive.createFile(KSDrive.FileType.KSheet)
console.log(url)
// 新建AP文件
url = KSDrive.createFile(KSDrive.FileType.AP)
console.log(url)
// 文件另存
url = KSDrive.createFile(KSDrive.FileType.KSheet, {
  source: 'https://www.kdocs.cn/l/cqQwuiG2mo7E',
  name: '复制表格'
})
console.log(url)
```

## openFile()

额外打开一个文件，并返回一个 JavaScript 对象File。

### 示例

```js
let file = KSDrive.openFile('https://www.kdocs.cn/l/xxxxxxxxxxxx')
console.log(file.Application.ActiveSheet.Range('A1').Text)
file.close()
```

### 参数

| 名称 | 类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| openInfo | URL / FileInfo | 是 | 打开文件的信息，可以为文件分享链接或者FileInfo |

### 返回值

File - 一个 JavaScript 对象

## listFiles()

列出某个目录下的所有文件和对应信息

### 示例

```js
// 遍历获取某个文件夹下的所有文件的文件名
for (let offset = 0; offset >= 0; ) {
  const list = KSDrive.listFiles({
    dirUrl: 'https://www.kdocs.cn/mine/xxxxxxxxxx',
    offset: offset,
    count: 100
  })
  for (let i = 0; i < list.files.length; i++) {
    console.log(list.files[i].fileName)
  }
  offset = list.nextOffset
}
```

### 参数

| 名称 | 类型 | 默认值 | 必填 | 说明 |
| --- | --- | --- | --- | --- |
| options | object | undefined | 否 | 一个 JavaScript 对象，undefined 时获取我的云文档目录下面的文件数据，详细参数如下所示 |

### 详细参数

| 参数名 | 参数类型 | 默认值 | 必填 | 说明 |
| --- | --- | --- | --- | --- |
| dirUrl | string |  | false | 目录链接，如`https://www.kdocs.cn/mine/xxxxxx`，为空时获取我的云文档目录下面的文件数据 |
| offset | number | 0 | false | 开始位置。通常由listFiles()函数返回。比如，listFiles()函数在某次检索中返回了 nextOffset 为 100，而想要获取更多文件信息，则下一次调用listFiles()函数时把 100 作为此可选参数传入。 |
| count | number | 30 | false | 文件个数 |
| includeExts | string[] |  | false | 指定文件类型,支持参数及对应关系，ksheet:"表格",et:"WPS 表格",db:"多维表",otl:"文档",wpp:"演示",wps:"WPS 文字" |

### 返回值

FilesInfo - 一个 JavaScript 对象，文件信息

## File

打开文件函数openFile()返回的一个 JavaScript 对象。

### 属性

| 名称 | 类型 | 说明 |
| --- | --- | --- |
| Application | [Application(ET/Ksheet/DBT)](../excel/workbook/Application.md) | 被打开文件的操作对象，目前支持 et,ksheet,dbt |
| close | Function | 关闭文件的函数，使用完 file 对象之后调用，关闭打开的文件 |

## FilesInfo

获取文件夹信息函数listFiles(options)返回的一个 JavaScript 对象。

### 属性

| 名称 | 类型 | 说明 |
| --- | --- | --- |
| files | FileInfo[] | 文件信息，详细参数如下所示 |
| nextOffset | number | 下一页的偏移量，可以作为listFiles(options)的参数而输出下一页文件内容，当下一页为空时，nextOffset 为-1 |

### FileInfo

| 名称 | 类型 | 说明 |
| --- | --- | --- |
| fileName | string | 文件名 |
| fileId | string | 加密后的文件 id |
| createTime | number | 文件创建时间戳 |
| updateTime | number | 文件修改时间戳 |
