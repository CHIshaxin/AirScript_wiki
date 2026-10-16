---
title: "WPS Docs API"
source: https://airsheet.wps.cn/docs/en-US/apiV2/advanced/KSDrive.html
updated: 1789700916000
hash: CDMDSL1T
---

# WPS Docs API

AirScript provides a global KSDrive object, through which you can easily **view, modify, and create** your cloud documents.

Tip

When using the KSDrive object to operate cloud documents, make sure you have added the `WPS Docs API` service. You can add it in the Services menu of the script editor.

### Quick Start

```js
// Open specified document
let file = KSDrive.openFile('https://www.wps.com/l/xxxxxxxxxxxx')
// Print content of cell A1 in specified document
console.log(file.Application.Range('A1').Text)
// Call close to close document after use, otherwise KSDrive.openFile cannot be called again
file.close()
// Get et, ksheet document list under My Cloud Documents
const fileList = KSDrive.listFiles({ includeExts: ['et', 'ksheet'] })
// Open first document under My Cloud Documents directory
file = KSDrive.openFile(fileList.files[0])
console.log(file.Application.Range('A1').Text)
// Close document
file.close()
```

### Property List

| Property Name | Data Type | Description |
|---|---|---|
| [FileType](#filetype) | object | Collection of supported file types |

### Method List

| Method Name | Return Type | Description |
|---|---|---|
| [createFile()](#createfile) | string | Create or save as a file |
| [openFile()](#openfile) | [File](#file) | Open an additional file |
| [listFiles()](#listfiles) | [FilesInfo](#filesinfo) | List spreadsheet files in a directory |

## FileType

File types supported by cloud documents, can be used to specify new file type when creating files.

### Property Description

| Property | Data Type | Description |
|---|---|---|
| AP | string | Smart Document |
| KSheet | string | Smart Spreadsheet |
| ET | string | Spreadsheet |
| DB | string | Multidimensional Table |

## createFile()

Create a new file, or save a source file as a new file.

### Parameters

| Name | Type | Required | Description |
|---|---|---|---|
| type | [FileType](#filetype) | Yes | Type of new file |
| createOptions | CreateOptions | Yes | Parameter options for new file |

### CreateOptions Object Description

| Name | Type | Required | Description |
|---|---|---|---|
| name | string | Yes | File name of new file |
| dirUrl | string | No | File directory of new file |
| source | string | No | Save target file as new file |

### Return Value

url - string URL of new file

### Example

```js
// Create ET file with specified save location
let url = KSDrive.createFile(KSDrive.FileType.ET, {
  name: 'et_test',
  dirUrl: 'specified save location'
})
console.log(url)
// Create new DB file
url = KSDrive.createFile(KSDrive.FileType.DB)
console.log(url)
// Create new KSheet file
url = KSDrive.createFile(KSDrive.FileType.KSheet)
console.log(url)
// Create new AP file
url = KSDrive.createFile(KSDrive.FileType.AP)
console.log(url)
// Save file as
url = KSDrive.createFile(KSDrive.FileType.KSheet, {
  source: 'https://www.wps.com/l/cqQwuiG2mo7E',
  name: 'Copy Spreadsheet'
})
console.log(url)
```

## openFile()

Open an additional file and return a JavaScript object [File](#file).

### Example

```js
let file = KSDrive.openFile('https://www.wps.com/l/xxxxxxxxxxxx')
console.log(file.Application.ActiveSheet.Range('A1').Text)
file.close()
```

### Parameters

| Name | Type | Required | Description |
|---|---|---|---|
| openInfo | URL / [FileInfo](#fileinfo) | Yes | Information for opening file, can be file share link or [FileInfo](#fileinfo) |

### Return Value

[File](#file) - A JavaScript object

## listFiles()

List all files and corresponding information in a directory.

### Example

```js
// Iterate to get all file names in a folder
for (let offset = 0; offset >= 0; ) {
  const list = KSDrive.listFiles({
    dirUrl: 'https://www.wps.com/mine/xxxxxxxxxx',
    offset: offset,
    count: 100
  })
  for (let i = 0; i < list.files.length; i++) {
    console.log(list.files[i].fileName)
  }
  offset = list.nextOffset
}
```

### Parameters

| Name | Type | Default | Required | Description |
|---|---|---|---|---|
| options | object | undefined | No | A JavaScript object, when undefined gets file data under My Cloud Documents directory, detailed parameters as shown below |

### Detailed Parameters

| Parameter | Type | Default | Required | Description |
|---|---|---|---|---|
| dirUrl | string |  | false | Directory link, such as `https://www.wps.com/mine/xxxxxx`, when empty gets file data under My Cloud Documents directory |
| offset | number | 0 | false | Start position. Usually returned by [listFiles()](#listfiles) function. For example, if [listFiles()](#listfiles) returns nextOffset as 100 in a retrieval, pass 100 as this optional parameter in the next call to get more file information. |
| count | number | 30 | false | Number of files |
| includeExts | string[] |  | false | Specify file types, supported parameters and mappings: ksheet:"Spreadsheet", et:"WPS Spreadsheet", db:"Multidimensional Table", otl:"Document", wpp:"Presentation", wps:"WPS Writer" |

### Return Value

[FilesInfo](#filesinfo) - A JavaScript object containing file information

## File

A JavaScript object returned by the [openFile()](#openfile) function.

### Properties

| Name | Type | Description |
|---|---|---|
| Application | [Application(ET/Ksheet/DBT)](/docs/en-US/api/excel/workbook/Application.html) | Operation object for opened file, currently supports et, ksheet, dbt |
| close | Function | Function to close file, call after using file object to close opened file |

## FilesInfo

A JavaScript object returned by the [listFiles(options)](#listfiles-options) function.

### Properties

| Name | Type | Description |
|---|---|---|
| files | [FileInfo](#fileinfo)[] | File information, detailed parameters as shown below |
| nextOffset | number | Offset for next page, can be used as parameter for [listFiles(options)](#listfiles-options) to output next page content, nextOffset is -1 when next page is empty |

### FileInfo

| Name | Type | Description |
|---|---|---|
| fileName | string | File name |
| fileId | string | Encrypted file id |
| createTime | number | File creation timestamp |
| updateTime | number | File modification timestamp |
