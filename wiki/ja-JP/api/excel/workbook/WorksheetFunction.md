---
title: "WorksheetFunction"
source: https://airsheet.wps.cn/docs/ja-JP/api/excel/workbook/WorksheetFunction.html
updated: 1789700916000
hash: BB9_CT-D
---

# WorksheetFunction

WorksheetFunction は、Sum、Average、Min、Max などの一般的な Excel 関数を公開する組み込みの Excel オブジェクトです。スクリプト内で WorksheetFunction を使用して、分析や計算のためにこれらの関数を呼び出します。

結果を変数に代入したり、数値処理、文字列処理、より複雑な分析のためにインラインで関数を呼び出したりできます。

たとえば、Sum は列の合計、Average は平均、Max は最大値、Min は最小値を返します。

#### メソッド

| メソッド | 戻り値の型 | 説明 |
|---|---|---|
| [Average()](#average) | Number | 範囲内の数値の平均 |
| [AverageIf()](#averageif) | Number | 条件を満たすセルの平均 |
| [Small()](#small) | Number | データ集合の k 番目に小さい値 |
| [Large()](#large) | Number | 配列または列の k 番目に大きい値 |
| [Min()](#min) | Number | 配列または列の最小値 |
| [Max()](#max) | Number | 配列または列の最大値 |
| [Sum()](#sum) | Number | 範囲内の数値の合計 |

## Average()

指定した範囲内の数値の平均を返します。

### パラメーター

| パラメーター | 型 | 既定値 | 必須 | 説明 |
|---|---|---|---|---|
| range | [Range](/docs/ja-JP/api/excel/workbook/Range.html) | null | はい | 平均を求める範囲 |

### 戻り値

Number - 範囲内の数値の平均

```js
// 列 A に名前、列 B にスコアがある表の場合、列 B の平均：

const range = Range('B:B')
const average = WorksheetFunction.Average(range)
console.log(average)
```

## AverageIf()

条件を満たすセルの平均を返します。

### パラメーター

| パラメーター | 型 | 既定値 | 必須 | 説明 |
|---|---|---|---|---|
| range | [Range](/docs/ja-JP/api/excel/workbook/Range.html) | null | はい | 評価する範囲 |
| criteria | String | null | はい | 条件。数値、式、参照、またはテキスト（例: 32、"32"、">32"、"apples"、または B4） |
| average_range | [Range](/docs/ja-JP/api/excel/workbook/Range.html) | null | いいえ | 実際に平均するセル。省略時は range が使用されます |

### 戻り値

Number - 一致する数値の平均

```js
// 列 A の名前が "Alice" のとき列 B のスコアの平均

const rangeA = Range('A:A')
const rangeB = Range('B:B')
const average = WorksheetFunction.AverageIf(rangeA, 'Alice', rangeB)
console.log(average)
```

## Small()

データ集合の k 番目に小さい値を返します。

配列が空の場合、Small() は `#NUM!` を返します。

k <= 0 または k がデータ点の数を超える場合、Small は `#NUM!` を返します。

データ点の数を n とすると、SMALL(array,1) は最小値、SMALL(array,n) は最大値です。

### パラメーター

| パラメーター | 型 | 既定値 | 必須 | 説明 |
|---|---|---|---|---|
| range | [Range](/docs/ja-JP/api/excel/workbook/Range.html) | null | はい | k 番目に小さい値を取得する配列または列 |
| k | Number | null | はい | 小さい方からの順位。k は 0 より大きく、要素数以下である必要があります |

### 戻り値

Number - k 番目に小さい値

### 例

```js
// 列 B のスコアで 3 番目に小さい値

const range = Range('B:B')
const third = WorksheetFunction.Small(range, 3)
console.log(third)
```

## Large()

配列または列の k 番目に大きい値を返します（例: 最高点、2 位、3 位）。

配列が空の場合、Large は `#NUM!` を返します。

k <= 0 または k がデータ点の数を超える場合、Large は `#NUM!` を返します。

データ点の数が n のとき、LARGE(array,1) は最大値、LARGE(array,n) は最小値です。

### パラメーター

| パラメーター | 型 | 既定値 | 必須 | 説明 |
|---|---|---|---|---|
| range | [Range](/docs/ja-JP/api/excel/workbook/Range.html) | null | はい | k 番目に大きい値を取得する配列または列 |
| k | Number | null | はい | 大きい方からの順位。k は 0 より大きく、要素数以下である必要があります |

### 戻り値

Number - k 番目に大きい値

### 例

```js
// 列 B のスコアで 4 番目に大きい値

const range = Range('B:B')
const fourth = WorksheetFunction.Large(range, 4)
console.log(fourth)
```

## Min()

配列または列の最小値を返します。

引数には数値、名前、配列、参照を指定できます。

引数リストに直接入力した論理値と数値の文字列表現は含まれます。

引数が配列または参照の場合、数値のみが使用されます。配列または参照内の空白セル、論理値、テキストは無視されます。

数値がない場合、Min は 0 を返します。

エラー、または数値に変換できないテキストがあるとエラーになります。

### パラメーター

| パラメーター | 型 | 既定値 | 必須 | 説明 |
|---|---|---|---|---|
| range | [Range](/docs/ja-JP/api/excel/workbook/Range.html) | null | はい | 最小値を取得する数値の範囲 |

### 戻り値

Number - 最小値

### 例

```js
// 列 B のスコアの最小値

const range = Range('B:B')
const min = WorksheetFunction.Min(range)
console.log(min)
```

## Max()

配列または列の最大値を返します。

引数には数値、名前、配列、参照を指定できます。

引数リストに直接入力した論理値と数値の文字列表現は含まれます。

引数が配列または参照の場合、数値のみが使用されます。配列または参照内の空白セル、論理値、テキストは無視されます。

数値がない場合、Max は 0 を返します。

エラー、または数値に変換できないテキストがあるとエラーになります。

### パラメーター

| パラメーター | 型 | 既定値 | 必須 | 説明 |
|---|---|---|---|---|
| range | [Range](/docs/ja-JP/api/excel/workbook/Range.html) | null | はい | 最大値を求める配列または列 |

### 戻り値

Number - 最大値

### 例

```js
// 列 B のスコアの最大値

const range = Range('B:B')
const max = WorksheetFunction.Max(range)
console.log(max)
```

## Sum()

範囲内の数値の合計を返します。

引数リストに直接入力した数値、論理値、数値の文字列表現は含まれます。

引数が配列または参照の場合、その配列または参照内の数値のみが合計されます。空のセル、論理値、テキストは無視されます。

エラー、または数値に変換できないテキストがあるとエラーになります。

### パラメーター

| パラメーター | 型 | 既定値 | 必須 | 説明 |
|---|---|---|---|---|
| Array | [Range](/docs/ja-JP/api/excel/workbook/Range.html) | null | はい | 合計する引数 |

### 戻り値

Number - すべての数値の合計

### 例

```js
// 列 B のスコアの合計

const range = Range('B:B')
const sum = WorksheetFunction.Sum(range)
console.log(sum)
```

## Match()

範囲内での値の相対位置を返します。

項目そのものではなく位置が必要なときに Lookup(Object, Object, Object) の代わりに Match を使用します。

Lookup_value は lookup_array 内で検索する値です（例: 最終的に電話番号が必要なときの電話帳内の名前）。

Lookup_value は数値、テキスト、論理値、またはそれらへの参照にできます。

match_type が 1 のとき、Match は lookup_value 以下で最大の値を検索します。lookup_array は昇順に並べ替えられている必要があります: ... -2, -1, 0, 1, 2, ..., A-Z, false, true。

match_type が 0 のとき、Match は最初に完全一致する値を検索します。lookup_array の順序は任意です。

match_type が -1 のとき、Match は lookup_value 以上で最小の値を検索します。lookup_array は降順に並べ替えられている必要があります: true, false, Z-A, ... 2, 1, 0, -1, -2, ...。match_type を省略すると 1 とみなされます。

Match は lookup_array 内の位置を返し、値そのものではありません。例: MATCH("b",{"a","b","c"},0) は 2 を返します。

テキストに対する Match の大文字／小文字の区別はありません。

一致がない場合、Match は #N/A を返します。

match_type が 0 で lookup_value がテキストのとき、ワイルドカード ? と * が使用できます。? は 1 文字、* は任意の連続に一致します。文字としての ? や * を検索するときは ~ を前置します。

### パラメーター

| パラメーター | 型 | 既定値 | 必須 | 説明 |
|---|---|---|---|---|
| Lookup_value | Object | null | はい | 検索する値 |
| Lookup_array | [Range](/docs/ja-JP/api/excel/workbook/Range.html) | null | はい | 値を含む可能性のある連続範囲。配列または参照である必要があります |
| Match_type | Object | null | いいえ | -1、0、または 1: Excel における lookup_value と lookup_array の比較方法 |

### 戻り値

Number - lookup_array 内における lookup_value の位置

### 例

```js
// 列 A の名前を使って "Alice" のスコアを取得

const range = Range('A:A')
const rowIndex = WorksheetFunction.Match('Alice', range, 0)
const score = Range(`B${rowIndex}`).Value
console.log(score)
```
