---
title: "WPS Docs API"
source: https://airsheet.wps.cn/docs/ja-JP/api/advanced/KSDrive.html
updated: 1789700916000
hash: BxjwgBZt
---

# WPS Docs API

AirScript はグローバルな KSDrive オブジェクトを公開します。これを使うと WPS Docs を**閲覧・編集・作成**できます。

ヒント

KSDrive を使う前に、スクリプトエディターの **サービス** メニューから `WPS Docs API` サービスを追加してください。

### クイックスタート

```js
// 指定ドキュメントを開く
let file = KSDrive.openFile('https://www.wps.com/l/xxxxxxxxxxxx')
// セル A1 を出力
console.log(file.Application.Range('A1').Text)
// 使用後は close を呼ぶ。close しないと KSDrive.openFile を再度使えない
file.close()
// マイ WPS Docs 配下の et と ksheet ファイルを一覧
const fileList = KSDrive.listFiles({ includeExts: ['et', 'ksheet'] })
// マイ WPS Docs の先頭ファイルを開く
file = KSDrive.openFile(fileList.files[0])
console.log(file.Application.Range('A1').Text)
// ドキュメントを閉じる
file.close()
```

### プロパティ

| プロパティ | 型 | 説明 |
|---|---|---|
| [FileType](#filetype) | object | 対応するファイル種別の定数 |

### メソッド

| メソッド | 戻り値の型 | 説明 |
|---|---|---|
| [createFile()](#createfile) | string | ファイルを作成する、または別名でコピーを保存する |
| [openFile()](#openfile-openinfo) | [File](#file) | 別ファイルを開く |
| [listFiles()](#listfiles) | [FilesInfo](#filesinfo) | フォルダー内のスプレッドシートファイルを一覧する |

## FileType

対応する WPS Docs のファイル種別。ファイル作成時に種別を指定するために使います。

### プロパティ

| プロパティ | 型 | 説明 |
|---|---|---|
| AP | string | スマートドキュメント |
| KSheet | string | スマートスプレッドシート |
| ET | string | スプレッドシート |
| DB | string | 多次元表 |

## createFile()

新規ファイルを作成するか、元ファイルを別名で保存します。

### パラメーター

| パラメーター | 型 | 必須 | 説明 |
|---|---|---|---|
| type | [FileType](#filetype) | はい | 新規ファイルの種別 |
| createOptions | CreateOptions | はい | 新規ファイルのオプション |

### CreateOptions

| パラメーター | 型 | 必須 | 説明 |
|---|---|---|---|
| name | string | はい | ファイル名 |
| dirUrl | string | いいえ | フォルダーの URL |
| source | string | いいえ | 別名保存の元ファイル URL |

### 戻り値

url - 新規ファイルの string URL

### 例

```js
// フォルダーに ET ファイルを作成
let url = KSDrive.createFile(KSDrive.FileType.ET, {
  name: 'et_test',
  dirUrl: 'https://www.wps.com/mine/folder_id'
})
console.log(url)
// DB ファイルを作成
url = KSDrive.createFile(KSDrive.FileType.DB)
console.log(url)
// KSheet ファイルを作成
url = KSDrive.createFile(KSDrive.FileType.KSheet)
console.log(url)
// AP ファイルを作成
url = KSDrive.createFile(KSDrive.FileType.AP)
console.log(url)
// 別名でコピー保存
url = KSDrive.createFile(KSDrive.FileType.KSheet, {
  source: 'https://www.wps.com/l/cqQwuiG2mo7E',
  name: 'Copy_spreadsheet'
})
console.log(url)
```

## openFile()

別ファイルを開き、JavaScript の [File](#file) オブジェクトを返します。

### 例

```js
let file = KSDrive.openFile('https://www.wps.com/l/xxxxxxxxxxxx')
console.log(file.Application.ActiveSheet.Range('A1').Text)
file.close()
```

### パラメーター

| パラメーター | 型 | 必須 | 説明 |
|---|---|---|---|
| openInfo | URL / [FileInfo](#fileinfo) | はい | 共有リンクまたは [FileInfo](#fileinfo) |

### 戻り値

[File](#file) - JavaScript オブジェクト

## listFiles()

フォルダー内のすべてのファイルとメタデータを一覧します。

### 例

```js
// フォルダー内の全ファイル名を反復
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

### パラメーター

| パラメーター | 型 | 既定値 | 必須 | 説明 |
|---|---|---|---|---|
| options | object | undefined | いいえ | オプションオブジェクト。`undefined` の場合はマイ WPS Docs 配下を一覧。詳細は後述。 |

### オプションのフィールド

| パラメーター | 型 | 既定値 | 必須 | 説明 |
|---|---|---|---|---|
| dirUrl | string |  | いいえ | フォルダーリンク。例: `https://www.wps.com/mine/xxxxxx`。空の場合はマイ WPS Docs。 |
| offset | number | 0 | いいえ | 開始オフセット。通常は [listFiles()](#listfiles) から。`nextOffset` が 100 なら、次ページ取得の [listFiles()](#listfiles) では 100 を渡す。 |
| count | number | 30 | いいえ | ページサイズ |
| includeExts | string[] |  | いいえ | ファイル種別フィルター: ksheet スマートスプレッドシート、et WPS スプレッドシート、db 多次元表、otl ドキュメント、wpp プレゼンテーション、wps WPS ライター |

### 戻り値

[FilesInfo](#filesinfo) - ファイル情報を含む JavaScript オブジェクト

## File

[openFile()](#openfile-openinfo) が返す JavaScript オブジェクト。

### プロパティ

| パラメーター | 型 | 説明 |
|---|---|---|
| Application | [Application(ET/Ksheet/DBT)](/docs/ja-JP/api/excel/workbook/Application.html) | 開いたファイル用 API。et、ksheet、dbt をサポート |
| close | Function | ファイルを閉じる。`file` オブジェクトの使用後に呼ぶ |

## FilesInfo

[listFiles(options)](#listfiles-options) が返す JavaScript オブジェクト。

### プロパティ

| パラメーター | 型 | 説明 |
|---|---|---|
| files | [FileInfo](#fileinfo)[] | ファイルエントリ。詳細は後述 |
| nextOffset | number | [listFiles(options)](#listfiles-options) の次ページ用オフセット。次ページがない場合は `-1` |

### FileInfo

| パラメーター | 型 | 説明 |
|---|---|---|
| fileName | string | ファイル名 |
| fileId | string | 暗号化されたファイル ID |
| createTime | number | 作成日時（タイムスタンプ） |
| updateTime | number | 更新日時（タイムスタンプ） |
