---
title: "Workbook"
source: https://airsheet.wps.cn/docs/ja-JP/api/excel/workbook/Workbook.html
updated: 1789700916000
hash: DRMvKkb5
---

# Workbook

Workbook（ブック）はファイルの作業領域であり、すべてのシートを含みます。

Workbook オブジェクトのプロパティとメソッドは、以下の一覧を参照してください。

#### プロパティ

| プロパティ | 型 | 説明 |
| --- | --- | --- |
| [ActiveSheet](#activesheet) | [Sheet](Sheet.md) | ブック内のアクティブなワークシート |
| [Sheets](#sheets) | [Sheets](Application.md#sheets) | すべてのワークシート オブジェクトのコレクション |
| [ReadOnly](#readonly) | Boolean | ドキュメントが読み取り専用かどうか |
| [ReadOnlyComment](#readonlycomment) | Boolean | ドキュメントがコメントのみの読み取りを許可するかどうか |
| [SupportReadOnlyComment](#supportreadonlycomment) | Boolean | ドキュメントがコメントのみの読み取りに対応するかどうか |

#### メソッド

| メソッド | 戻り値の型 | 説明 |
| --- | --- | --- |
| [Save()](#save) | String(JSON) | 変更をファイルに保存します |
| [GetComments()](#getcomments) | String(JSON) | ブック全体のコメントを取得します |
| [ExportAsFixedFormat()](#exportasfixedformat) | String(JSON) | シート全体を PDF または画像としてエクスポートします |

## ActiveSheet

ブック内のアクティブなワークシートです。

`Application.ActiveSheet` を使用しても同じです。

#### データ型

[Sheet](Sheet.md) - アクティブなブックのアクティブなワークシート

#### 例

```js
// 次の 2 つの形式は同等です
console.log(Application.ActiveWorkbook.ActiveSheet.Name) // Sheet1
console.log(Application.ActiveSheet.Name) // Sheet1
```

## Sheets

現在のファイルで操作できるすべてのシートを [Sheets](Application.md#sheets) オブジェクトとして取得します。

#### データ型

[Sheets](Application.md#sheets)

#### 例

```js
// 次の 2 つの形式は同等です
let sheets = Application.ActiveWorkbook.Sheets
sheets = Application.Sheets
// すべてのワークシート名を出力
for (let i = 1; i <= sheets.Count; i++) {
    console.log((sheets.Item(i).Name))
}
```

## ReadOnly

ドキュメントが読み取り専用かどうかを返します。このプロパティは読み取り専用です。

#### データ型

Boolean - ドキュメントが読み取り専用かどうか

#### 例

```js
// 現在のドキュメントが読み取り専用かどうかを出力
console.log(Application.ActiveWorkbook.ReadOnly) //false
```

## ReadOnlyComment

ドキュメントにコメントのみの読み取り権限があるかどうかを返します。このプロパティは読み取り専用です。

#### データ型

Boolean - ドキュメントにコメントのみの読み取り権限があるかどうか

#### 例

```js
// 現在のドキュメントがコメントのみ読み取り可能かどうかを出力
console.log(Application.ActiveWorkbook.ReadOnlyComment) //true
```

## Save()

変更をファイルに保存します。

#### 戻り値の型

String(JSON) - 保存結果。形式は次のとおりです。

| プロパティ | 型 | 説明 |
| --- | --- | --- |
| result | String | 保存ステータス |
| size | Number | ファイルサイズ（バイト） |
| version | Number | バージョン |

**保存ステータスの値**:

| ステータス | 説明 |
| --- | --- |
| ok | バージョンが正常に保存されました。バージョン履歴で確認できます |
| nochange | ドキュメントに変更がありません。保存するバージョンはありません |
| SavedEmptyFile | 空ファイルの保存はサポートされていません。カーネル保存後にファイルが空のときに発生します |
| SpaceFull | ストレージがいっぱいです |
| QueneFull | 保存の頻度が高すぎます。サーバーの保存キューがいっぱいでリクエストがキューに入ったときに発生します |
| fail | 保存に失敗しました |

#### 例

```js
// 変更を保存
let res = Application.ActiveWorkbook.Save()
if (res.result === 'ok') {
  console.log('成功')
} else {
  console.error('失敗:', res.result)
}
```

## GetComments()

ブック全体のコメントを取得します。

#### 戻り値の型

String(JSON) - ブックのコメント コレクション。形式は次のとおりです。

| プロパティ | 型 | 説明 |
| --- | --- | --- |
| CellComments | `Array<Object>` | コメント項目 |
| PosInfo | `String` | セル情報 |
| SheetName | `String` | シート名 |
| UserIds | `Array<String>` | ユーザー ID のコレクション |

**コメント項目**:

| プロパティ | 型 | 説明 |
| --- | --- | --- |
| DateTime | `String` | タイムスタンプ |
| Text | `String` | コメント本文 |
| Time | `String` | 書式設定済みの時刻 |
| UserId | `String` | ユーザー ID |

#### 例

```js
// ブック全体のコメントを取得
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

シート全体を PDF または画像としてエクスポートし、エクスポート先の URL を返します。

#### パラメーター

| パラメーター | 型 | 既定値 | 必須 | 説明 |
| --- | --- | --- | --- | --- |
| Type | Enum |  | 任意 | エクスポートの種類。[Enum.XlFixedFormatType](Enum.md#xlfixedformattype) を参照。画像と PDF のエクスポートのみサポート |

#### 戻り値の型

String(JSON) - エクスポート URL を含む JSON 文字列

#### 例 1: PDF をエクスポート

```js
// シート全体をエクスポート
let json = Application.ActiveWorkbook.ExportAsFixedFormat()
console.log(json.url)
```

#### 例 2: 画像をエクスポート

```js
// シート全体をエクスポート
let json = Application.ActiveWorkbook.ExportAsFixedFormat({
  Type: Application.Enum.XlFixedFormatType.xlTypeIMG
})
console.log(json.url)
```
