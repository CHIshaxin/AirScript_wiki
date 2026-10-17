---
title: "WPS Docs API"
source: https://airsheet.wps.cn/docs/en-US/api/advanced/KSDrive.html
updated: 1789700916000
hash: D92-cCUr
---

# WPS Docs API

AirScript exposes a global KSDrive object. With it you can **view, edit, and create** your WPS Docs.

Tip

Before using KSDrive, add the `WPS Docs API` service from the script editor **Services** menu.

### Quick start

```js
// Opens the specified document
let file = KSDrive.openFile('https://www.wps.com/l/xxxxxxxxxxxx')
// Prints cell A1
console.log(file.Application.Range('A1').Text)
// Call close when done; otherwise KSDrive.openFile cannot be used again
file.close()
// Lists et and ksheet files under My WPS Docs
const fileList = KSDrive.listFiles({ includeExts: ['et', 'ksheet'] })
// Opens the first file in My WPS Docs
file = KSDrive.openFile(fileList.files[0])
console.log(file.Application.Range('A1').Text)
// Closes the document
file.close()
```

### Properties

| Property | Type | Description |
| --- | --- | --- |
| [FileType](#filetype) | object | Supported file type constants |

### Methods

| Method | Return type | Description |
| --- | --- | --- |
| [createFile()](#createfile) | string | Creates a file or saves a copy as new |
| [openFile()](#openfile-openinfo) | [File](#file) | Opens another file |
| [listFiles()](#listfiles) | [FilesInfo](#filesinfo) | Lists spreadsheet files in a folder |

## FileType

Supported WPS Docs file types; use when creating a file to set its type.

### Properties

| Property | Type | Description |
| --- | --- | --- |
| AP | string | Smart document |
| KSheet | string | Smart Spreadsheet |
| ET | string | Spreadsheet |
| DB | string | Multi-dimensional table |

## createFile()

Creates a new file, or saves a source file as a new file.

### Parameters

| Name | Type | Required | Description |
| --- | --- | --- | --- |
| type | [FileType](#filetype) | Yes | Type of the new file |
| createOptions | CreateOptions | Yes | Options for the new file |

### CreateOptions

| Name | Type | Required | Description |
| --- | --- | --- | --- |
| name | string | Yes | File name |
| dirUrl | string | No | Folder URL |
| source | string | No | Source file URL to save as a new file |

### Return value

url - string URL of the new file

### Example

```js
// Creates an ET file at a folder
let url = KSDrive.createFile(KSDrive.FileType.ET, {
  name: 'et_test',
  dirUrl: 'https://www.wps.com/mine/folder_id'
})
console.log(url)
// Creates a DB file
url = KSDrive.createFile(KSDrive.FileType.DB)
console.log(url)
// Creates a KSheet file
url = KSDrive.createFile(KSDrive.FileType.KSheet)
console.log(url)
// Creates an AP file
url = KSDrive.createFile(KSDrive.FileType.AP)
console.log(url)
// Save as copy
url = KSDrive.createFile(KSDrive.FileType.KSheet, {
  source: 'https://www.wps.com/l/cqQwuiG2mo7E',
  name: 'Copy_spreadsheet'
})
console.log(url)
```

## openFile()

Opens another file and returns a JavaScript [File](#file) object.

### Example

```js
let file = KSDrive.openFile('https://www.wps.com/l/xxxxxxxxxxxx')
console.log(file.Application.ActiveSheet.Range('A1').Text)
file.close()
```

### Parameters

| Name | Type | Required | Description |
| --- | --- | --- | --- |
| openInfo | URL / [FileInfo](#fileinfo) | Yes | Share link or [FileInfo](#fileinfo) |

### Return value

[File](#file) - A JavaScript object

## listFiles()

Lists all files and metadata in a folder.

### Example

```js
// Iterates all file names in a folder
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
| --- | --- | --- | --- | --- |
| options | object | undefined | No | Options object; `undefined` lists files under My WPS Docs. Details below. |

### Options fields

| Name | Type | Default | Required | Description |
| --- | --- | --- | --- | --- |
| dirUrl | string |  | false | Folder link, e.g. `https://www.wps.com/mine/xxxxxx`. Empty means My WPS Docs. |
| offset | number | 0 | false | Start offset, usually from [listFiles()](#listfiles). If `nextOffset` was 100, pass 100 on the next [listFiles()](#listfiles) call for the next page. |
| count | number | 30 | false | Page size |
| includeExts | string[] |  | false | File type filter: ksheet Smart Spreadsheet, et WPS Spreadsheet, db multi-dimensional table, otl document, wpp presentation, wps WPS Writer |

### Return value

[FilesInfo](#filesinfo) - A JavaScript object with file information

## File

JavaScript object returned by [openFile()](#openfile-openinfo).

### Properties

| Name | Type | Description |
| --- | --- | --- |
| Application | [Application(ET/Ksheet/DBT)](../excel/workbook/Application.md) | API for the opened file; supports et, ksheet, dbt |
| close | Function | Closes the file; call when done with the `file` object |

## FilesInfo

JavaScript object returned by [listFiles(options)](#listfiles-options).

### Properties

| Name | Type | Description |
| --- | --- | --- |
| files | [FileInfo](#fileinfo)[] | File entries; details below |
| nextOffset | number | Next page offset for [listFiles(options)](#listfiles-options); `-1` when there is no next page |

### FileInfo

| Name | Type | Description |
| --- | --- | --- |
| fileName | string | File name |
| fileId | string | Encrypted file id |
| createTime | number | Created time (timestamp) |
| updateTime | number | Updated time (timestamp) |
