---
title: "WPS Docs API"
source: https://airsheet.wps.cn/docs/ja-JP/apiV2/advanced/KSDrive.html
updated: 1789700916000
hash: Z1loVnLW
---

# WPS Docs API

AirScript はグローバルな KSDrive オブジェクトを提供しており、クラウドドキュメントを**閲覧・編集・作成**できます。

ヒント

KSDrive オブジェクトでクラウドドキュメントを操作する場合、`WPS Docs API` サービスを追加済みであることを確認してください。スクリプトエディターの［サービス］メニューから追加できます。

### クイックスタート

```js
// 指定ドキュメントを開く
let file = KSDrive.openFile('https://www.wps.com/l/xxxxxxxxxxxx')
// 指定ドキュメントのセル A1 の内容を出力
console.log(file.Application.Range('A1').Text)
// 使用後は close を呼び出してドキュメントを閉じる。閉じないと KSDrive.openFile を再度呼べない
file.close()
// マイクラウドドキュメント配下の et、ksheet ドキュメント一覧を取得
const fileList = KSDrive.listFiles({ includeExts: ['et', 'ksheet'] })
// マイクラウドドキュメントディレクトリの先頭ドキュメントを開く
file = KSDrive.openFile(fileList.files[0])
console.log(file.Application.Range('A1').Text)
// ドキュメントを閉じる
file.close()
```

### プロパティ一覧

| プロパティ名 | データ型 | 説明 |
| --- | --- | --- |
| [FileType](#filetype) | object | サポートするファイルタイプの集合 |

### メソッド一覧

| メソッド名 | 戻り値の型 | 説明 |
| --- | --- | --- |
| [createFile()](#createfile) | string | ファイルを新規作成または別名保存 |
| [openFile()](#openfile-openinfo) | [File](#file) | 追加でファイルを開く |
| [listFiles()](#listfiles-options) | [FilesInfo](#filesinfo) | ディレクトリ内のスプレッドシートファイルを一覧 |

## FileType

クラウドドキュメントでサポートされるファイルタイプ。ファイル作成時に新規ファイルのタイプを指定するのに使います。

### プロパティの説明

| プロパティ | データ型 | 説明 |
| --- | --- | --- |
| AP | string | スマートドキュメント |
| KSheet | string | スマートスプレッドシート |
| ET | string | スプレッドシート |
| DB | string | 多次元表 |

## createFile()

新規ファイルを作成するか、元ファイルを新しいファイルとして保存します。

### パラメーター

| パラメーター | 型 | 必須 | 説明 |
| --- | --- | --- | --- |
| type | [FileType](#filetype) | はい | 新規ファイルのタイプ |
| createOptions | CreateOptions | はい | 新規ファイルのパラメーターオプション |

### CreateOptions オブジェクトの説明

| パラメーター | 型 | 必須 | 説明 |
| --- | --- | --- | --- |
| name | string | はい | 新規ファイルのファイル名 |
| dirUrl | string | いいえ | 新規ファイルの保存ディレクトリ |
| source | string | いいえ | 対象ファイルを新しいファイルとして保存 |

### 戻り値

url — 新規ファイルの string URL

### 例

```js
// 保存先を指定して ET ファイルを作成
let url = KSDrive.createFile(KSDrive.FileType.ET, {
  name: 'et_test',
  dirUrl: 'specified save location'
})
console.log(url)
// 新規 DB ファイルを作成
url = KSDrive.createFile(KSDrive.FileType.DB)
console.log(url)
// 新規 KSheet ファイルを作成
url = KSDrive.createFile(KSDrive.FileType.KSheet)
console.log(url)
// 新規 AP ファイルを作成
url = KSDrive.createFile(KSDrive.FileType.AP)
console.log(url)
// 名前を付けて保存
url = KSDrive.createFile(KSDrive.FileType.KSheet, {
  source: 'https://www.wps.com/l/cqQwuiG2mo7E',
  name: 'Copy Spreadsheet'
})
console.log(url)
```

## openFile()

追加でファイルを開き、JavaScript オブジェクト [File](#file) を返します。

### 例

```js
let file = KSDrive.openFile('https://www.wps.com/l/xxxxxxxxxxxx')
console.log(file.Application.ActiveSheet.Range('A1').Text)
file.close()
```

### パラメーター

| パラメーター | 型 | 必須 | 説明 |
| --- | --- | --- | --- |
| openInfo | URL / [FileInfo](#fileinfo) | はい | 開くファイルの情報。ファイル共有リンクまたは [FileInfo](#fileinfo) |

### 戻り値

[File](#file) — JavaScript オブジェクト

## listFiles()

ディレクトリ内のすべてのファイルと対応する情報を一覧します。

### 例

```js
// フォルダ内のすべてのファイル名を反復取得
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
| --- | --- | --- | --- | --- |
| options | object | undefined | いいえ | JavaScript オブジェクト。undefined のときはマイクラウドドキュメント直下のファイルデータを取得。詳細は下記 |

### 詳細パラメーター

| パラメーター | 型 | 既定値 | 必須 | 説明 |
| --- | --- | --- | --- | --- |
| dirUrl | string |  | いいえ | ディレクトリリンク。例：`https://www.wps.com/mine/xxxxxx`。空のときはマイクラウドドキュメント直下のファイルデータを取得 |
| offset | number | 0 | いいえ | 開始位置。通常は [listFiles()](#listfiles-options) の戻り値。例：[listFiles()](#listfiles-options) の取得で nextOffset が 100 の場合、次の呼び出しでこの任意パラメーターに 100 を渡すと、続きのファイル情報を取得できる。 |
| count | number | 30 | いいえ | 取得するファイル数 |
| includeExts | string[] |  | いいえ | ファイルタイプを指定。サポートするパラメーターと対応：ksheet:「スプレッドシート」、et:「WPS スプレッドシート」、db:「多次元表」、otl:「ドキュメント」、wpp:「プレゼンテーション」、wps:「WPS ライター」 |

### 戻り値

[FilesInfo](#filesinfo) — ファイル情報を含む JavaScript オブジェクト

## File

[openFile()](#openfile-openinfo) が返す JavaScript オブジェクト。

### プロパティ

| プロパティ名 | 型 | 説明 |
| --- | --- | --- |
| Application | [Application(ET/Ksheet/DBT)](../../api/excel/workbook/Application.md) | 開いたファイルの操作オブジェクト。現在 et、ksheet、dbt をサポート |
| close | Function | ファイルを閉じる関数。ファイルオブジェクト使用後に呼び出して開いたファイルを閉じる |

## FilesInfo

[listFiles(options)](#listfiles-options) が返す JavaScript オブジェクト。

### プロパティ

| プロパティ名 | 型 | 説明 |
| --- | --- | --- |
| files | [FileInfo](#fileinfo)[] | ファイル情報。詳細は下記 |
| nextOffset | number | 次ページ用オフセット。[listFiles(options)](#listfiles-options) のパラメーターとして次ページを出力するのに使える。次ページがないとき nextOffset は -1 |

### FileInfo

| プロパティ名 | 型 | 説明 |
| --- | --- | --- |
| fileName | string | ファイル名 |
| fileId | string | 暗号化されたファイル ID |
| createTime | number | ファイル作成時刻（タイムスタンプ） |
| updateTime | number | ファイル更新時刻（タイムスタンプ） |
