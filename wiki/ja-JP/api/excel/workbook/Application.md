---
title: "Application"
source: https://airsheet.wps.cn/docs/ja-JP/api/excel/workbook/Application.html
updated: 1789700916000
hash: a6xHwmba
---

# Application

ドキュメント操作の最上位オブジェクトです。ドキュメントの操作は、常にこのオブジェクトを直接または間接的に通じて行われます。

Application はファイルの最上位オブジェクトであり、新しいファイルを開いた場合も Application が返されます。

スクリプトでは、`Application` は現在のファイルの最上位オブジェクトを指します。常に 1 つだけ存在します。

Application オブジェクトのプロパティとメソッドは、以下の一覧を参照してください。

#### プロパティ

| プロパティ | 型 | 説明 |
|---|---|---|
| [ActiveSheet](#activesheet) | [Sheet](/docs/ja-JP/api/excel/workbook/Sheet.html) | アクティブなワークシート |
| [Sheets](#sheets) | [Sheets](/docs/ja-JP/api/excel/workbook/Application.html#sheets) | 現在のファイル内のすべてのワークシート |
| [ActiveWorkbook](#activeworkbook) | [Workbook](/docs/ja-JP/api/excel/workbook/Workbook.html) | 現在のドキュメント |
| [Selection](#selection) | [Range](/docs/ja-JP/api/excel/workbook/Range.html) | 現在の選択範囲 |
| [Cells](#cells) | [Range](/docs/ja-JP/api/excel/workbook/Range.html) | アクティブシート上のすべてのセル |
| [Columns](#columns) | [Range](/docs/ja-JP/api/excel/workbook/Range.html) | アクティブシート上のすべての列 |
| [Rows](#rows) | [Range](/docs/ja-JP/api/excel/workbook/Range.html) | アクティブシート上のすべての行 |
| [FileInfo](#fileinfo) | オブジェクト | 現在のドキュメントに関する情報 |
| [UserInfo](#userinfo) | オブジェクト | 現在のドキュメントのユーザー情報 |
| [Enum](#enum) | [Enum](/docs/ja-JP/api/excel/workbook/Enum.html) | すべての列挙型 |

#### メソッド

| メソッド | 戻り値 | 説明 |
|---|---|---|
| [Range(address)](#range) | [Range](/docs/ja-JP/api/excel/workbook/Range.html) | 現在の ActiveSheet 上の領域を取得します（address で領域を指定） |
| [Sheets(name)](#sheets-1) | [Sheet](/docs/ja-JP/api/excel/workbook/Sheet.html) | 名前が name のワークシートを取得します |

## ActiveSheet

アクティブなワークシートです。Sheet.Activate() でアクティブシートを切り替えられます。このプロパティは [Sheet](/docs/ja-JP/api/excel/workbook/Sheet.html#ActiveSheet) オブジェクトを返すため、アクティブなワークシートを操作できます。

スクリプトはサーバー上で隔離された環境で実行されるため、スクリプト環境の ActiveSheet がユーザーの ActiveSheet と一致しない場合があります。

ルール:

1. スクリプト実行時、スクリプト環境の ActiveSheet は、ユーザーの現在の ActiveSheet に合わせて切り替わります。
2. 関数によりスクリプト環境の ActiveSheet を変更しても、ユーザーの ActiveSheet は同期して変わりません。

#### データ型

[Sheet](/docs/ja-JP/api/excel/workbook/Sheet.html) - アクティブなワークシート

#### 例

```js
console.log(Application.ActiveSheet.Name) // Sheet1

// 名前が Sheet2 のワークシートに切り替え
Application.Sheets.Item('Sheet2').Activate()
console.log(Application.ActiveSheet.Name) // Sheet2

// アクティブシートのセル A1 に foo を設定
Application.ActiveSheet.Range('A1').Value = 'foo'
```

## Sheets

現在のファイルで操作可能なすべてのワークシートを [Sheets](/docs/ja-JP/api/excel/workbook/Application.html#sheets) オブジェクトとして取得します。

#### データ型

[Sheets](/docs/ja-JP/api/excel/workbook/Application.html#sheets)

#### 例

```js
// ブック（Workbook）内のすべてのワークシート（Sheet）のコレクション。以下の 2 形は同等
let sheets = Application.ActiveWorkbook.Sheets
sheets = Application.Sheets

// すべてのワークシート名を出力
for (let i = 1; i <= sheets.Count; i++) {
  console.log(sheets.Item(i).Name)
}
```

### Sheets.Count

ワークシートの数です。

#### データ型

数値 - ブック内のワークシート数

#### 例

```js
// 以下の 2 形は同等
let sheets = Application.ActiveWorkbook.Sheets
sheets = Application.Sheets

// ワークシート数を出力
console.log(sheets.Count) //1
```

### Sheets.DefaultNewSheetName

新しいワークシートの既定の名前です。

#### 戻り値

文字列 - ワークシート追加時に名前を指定しない場合、この名前を新しいシート名として使えます。

#### 例

```js
const defaultName = Application.Sheets.DefaultNewSheetName
// Worksheet オブジェクト
Application.Sheets.Add(
  null,
  Application.ActiveSheet.Name,
  1,
  Application.Enum.XlSheetType.xlWorksheet,
  defaultName
)
```

### Sheets.Add()

ワークシートを追加します。Before と After の両方を指定した場合、Before が優先されます。

#### パラメーター

| パラメーター | 型 | 既定値 | 必須 | 説明 |
|---|---|---|---|---|
| Before | 文字列/数値 |  | いいえ | After が空のとき必須。既存シートのインデックスまたは名前。新しいシートはこのシートの前に挿入されます |
| After | 文字列/数値 |  | いいえ | Before が空のとき必須。既存シートのインデックスまたは名前。新しいシートはこのシートの後に挿入されます |
| Count | 数値 | 1 | いいえ | 追加するシート数。既定は選択されているシートの数 |
| Type | Enum |  | いいえ | ワークシートの種類。[Enum.XlSheetType](/docs/ja-JP/api/excel/workbook/Enum.html#xlsheettype) を参照 |
| Name | Name |  | いいえ | ワークシート名 |

#### 例

```js
// ワークシートを追加
Application.Sheets.Add(
  null,
  Application.ActiveSheet.Name,
  1,
  Application.Enum.XlSheetType.xlWorksheet,
  'NewSheet'
)
```

### Sheets.Item()

名前またはインデックスでシートを選択します。

#### パラメーター

| パラメーター | 型 | 既定値 | 必須 | 説明 |
|---|---|---|---|---|
| index | 文字列/数値 |  | はい | 選択するシートの名前またはインデックス |

#### 戻り値

[Sheet](/docs/ja-JP/api/excel/workbook/Sheet.html) - 指定した名前のワークシート

#### 例

```js
// 名前が "Sheet2" のワークシートをアクティブ化
Application.Sheets.Item('Sheet2').Activate()

// インデックス 1 のワークシートをアクティブ化
Application.Sheets.Item(1).Activate()
```

### Sheets.Each()

すべてのシートを反復し、コールバックを実行します。

#### パラメーター

| パラメーター | 型 | 既定値 | 必須 | 説明 |
|---|---|---|---|---|
| callback | 関数 | null | はい | JS の配列 forEach に類似 |

#### 例

```js
// すべてのワークシート名を出力
Application.Sheets.Each(function (item) {
  console.log(item.Name) //Sheet1 Sheet2
})
```

## ActiveWorkbook

アクティブなファイルです。[Workbook](/docs/ja-JP/api/excel/workbook/Workbook.html) オブジェクトを返します。

#### データ型

[Workbook](/docs/ja-JP/api/excel/workbook/Workbook.html) - アクティブなファイル

#### 例

```js
// 保存
Application.ActiveWorkbook.Save()
```

## Selection

現在の選択範囲です。[Range](/docs/ja-JP/api/excel/workbook/Range.html) オブジェクトを返します。

#### データ型

[Range](/docs/ja-JP/api/excel/workbook/Range.html) - 現在の選択範囲

#### 例

```js
// ユーザーがスクリプト実行時の選択アドレスを出力（ここでは A1 を選択した想定）
console.log(Application.Selection.Address()) // $A$1
```

## Cells

現在のシート上のすべてのセルです。[Range](/docs/ja-JP/api/excel/workbook/Range.html) オブジェクトを返します。

#### データ型

[Range](/docs/ja-JP/api/excel/workbook/Range.html) - 現在のシート上のすべてのセル

#### 例

```js
const cells = Application.ActiveSheet.Cells
// 範囲内のセル数を出力
console.log(cells.Count) // 17179869184
```

## Columns

現在のシート上のすべての列です。[Range](/docs/ja-JP/api/excel/workbook/Range.html) オブジェクトを返します。

#### データ型

[Range](/docs/ja-JP/api/excel/workbook/Range.html) - 現在のシート上のすべての列

#### 例

```js
const columns = Application.ActiveSheet.Columns
// 列数を出力
console.log(columns.Count) // 16384
```

## Rows

現在のシート上のすべての行です。[Range](/docs/ja-JP/api/excel/workbook/Range.html) オブジェクトを返します。

#### データ型

[Range](/docs/ja-JP/api/excel/workbook/Range.html) - 現在のシート上のすべての行

#### 例

```js
let rows = Application.ActiveSheet.Rows

// 行数を出力
console.log(rows.Count) // 1048576
```

## FileInfo

現在のファイルに関する基本情報を返します。

#### データ型

オブジェクト - ファイル情報

| 名前 | 型 | 説明 |
|---|---|---|
| id | 文字列 | ファイル ID |
| name | 文字列 | ファイル名 |
| officeType | 文字列 | ドキュメントの種類 |
| creator | [CreatorObject](#creatorobject) | ドキュメント作成者 |
| size | 数値 | ファイルサイズ |
| groupId | 文字列 | ファイルのグループ ID |
| docType | 数値 | ドキュメントの種類（数値） |

#### CreatorObject

| 名前 | 型 | 説明 |
|---|---|---|
| id | 文字列 | 作成者 ID |
| name | 文字列 | 作成者名 |
| avatar_url | 文字列 | 作成者のアバター |
| logined | 真偽値 | ログイン済みかどうか |
| attrs | オブジェクト | プロパティオブジェクト |
| real_id | 文字列 | 実 ID |

#### 例

```javascript
// ファイル情報を出力
console.log(Application.FileInfo)
/*{
 "id": "<open_id>",
 ...
}*/
```

## UserInfo

現在のファイルに関するユーザー情報を返します。

#### データ型

オブジェクト - 現在のファイルのユーザー情報

| 名前 | 型 | 説明 |
|---|---|---|
| id | 文字列 | ユーザー ID |
| name | 文字列 | ユーザー名 |

#### 例

```javascript
// ユーザー情報を出力
console.log(Application.UserInfo)
```

## Enum

Application 配下に格納される列挙型です。

Application.Enum 経由で使用します。

#### データ型

[Enum](/docs/ja-JP/api/excel/workbook/Enum.html) - すべての列挙型

#### 例

```js
// ワークシート種別の列挙を出力
console.log(Application.Enum.XlSheetType)
//{"xlChart":-4109,"xlDialogSheet":-4116,"xlExcel4IntlMacroSheet":4,"xlExcel4MacroSheet":3,"xlWorksheet":-4167}
```

## Range()

現在の ActiveSheet 上の領域に対する [Range](/docs/ja-JP/api/excel/workbook/Range.html) を取得します。

#### パラメーター

| パラメーター | 型 | 必須 | 説明 |
|---|---|---|---|
| address | 文字列 | はい | アドレス。セルアドレスまたは範囲。例: "A1" または "A1:A10" |

#### 戻り値

[Range](/docs/ja-JP/api/excel/workbook/Range.html) - 指定アドレスの Range オブジェクト

#### 例

```js
// 以下の 2 形は同等。アクティブシートの A1 のテキストを取得
console.log(Application.Range('A1').Text) // A1 の内容

console.log(Application.ActiveSheet.Range('A1').Text) // A1 の内容
```

## Sheets()

関数として使用します。Sheets.Item() と同等です。[Sheet](/docs/ja-JP/api/excel/workbook/Sheet.html) オブジェクトを返します。

#### パラメーター

| パラメーター | 型 | 必須 | 説明 |
|---|---|---|---|
| name | 文字列 | はい | ワークシート名 |

#### 戻り値

[Sheet](/docs/ja-JP/api/excel/workbook/Sheet.html) - 指定した名前の Sheet オブジェクト

#### 例

```js
console.log(Application.Sheets.Count) // 1

// 以下の 2 形は同等
console.log(Application.Sheets('Sheet2').Range('A1').Text) // Sheet2 の A1 の内容
console.log(Application.Sheets.Item('Sheet2').Range('A1').Text) // Sheet2 の A1 の内容
```
