---
title: "AutoFilter"
source: https://airsheet.wps.cn/docs/ja-JP/api/excel/workbook/AutoFilter.html
updated: 1789700916000
hash: CD1JXhU3
---

# AutoFilter

ワークシートの AutoFilter です。データのフィルタリングや値の素早い検索に使います。1 列以上の列にフィルターを組み合わせられます。

フィルターにより、表示する行と隠す行を制御できます。1 列以上の列の値が条件を満たさない場合、行全体が非表示になります。数値や文字列でフィルターしたり、セルの塗りつぶしやフォントの色でフィルターしたりできます。

開発者は [Filters](#filters) プロパティで列フィルターのコレクションを取得します。[Range](#range) プロパティは、フィルターがかかった領域全体の [Range](Range.md) を返します。

注意

シートに AutoFilter を作成するには、ツールバーで手動によりフィルターを有効にするか、[Range](Range.md) の [AutoFilter](Range.md#autofilter) メソッドで領域の AutoFilter を有効にしてください。

AutoFilter オブジェクトのプロパティとメソッドは、以下の一覧を参照してください。

#### プロパティ

| プロパティ | 型 | 説明 |
| --- | --- | --- |
| [Filters](#filters) | [Filters](#filters) | フィルターのコレクション |
| [Range](#range) | [Range](Range.md) | フィルター後の範囲 |

#### メソッド

| メソッド | 戻り値 | 説明 |
| --- | --- | --- |
| [ApplyFilter()](#applyfilter) | undefined | 現在のシートにフィルターを適用する |
| [ShowAllData()](#showalldata) | undefined | すべてのフィルターを解除し、すべての行を表示する |

#### 使用例

ワークシートでの典型的な AutoFilter のシナリオ:

次のようなシートがあり、必要な行だけに絞り込みたいとします。

1. 名前が `Kim` の行にフィルター

```js
// AutoFilter オブジェクト
const autoFilter = ActiveSheet.AutoFilter
// 名前列のフィルター
const filter2 = autoFilter.Filters.Item(2)
// 値によるフィルター
filter2.Operator = Enum.XlAutoFilterOperator.xlFilterValues
filter2.Criteria1 = ['Kim']
autoFilter.ApplyFilter()
```

2. 国語の得点が上位 20% の行にフィルター

```js
const autoFilter = ActiveSheet.AutoFilter
const filterItem = autoFilter.Filters.Item(5)
filterItem.Operator = Enum.XlAutoFilterOperator.xlTop10Percent
filterItem.Criteria1 = '20'
autoFilter.ApplyFilter()
```

3. 数学の得点が上位 10 件の行にフィルター

```js
const autoFilter = ActiveSheet.AutoFilter
const filterItem = autoFilter.Filters.Item(6)
filterItem.Operator = Enum.XlAutoFilterOperator.xlTop10Items
filterItem.Criteria1 = '10'
autoFilter.ApplyFilter()
```

4. 名前列でセルの塗りつぶしが赤の行にフィルター

```js
const autoFilter = ActiveSheet.AutoFilter
const filterItem = autoFilter.Filters.Item(2)
filterItem.Operator = Enum.XlAutoFilterOperator.xlFilterCellColor
filterItem.Criteria1 = '#FF0000'
autoFilter.ApplyFilter()
```

5. 国語の得点が 100 以上 200 以下の行にフィルター

```js
const autoFilter = ActiveSheet.AutoFilter
const filterItem = autoFilter.Filters.Item(5)
filterItem.Operator = Enum.XlAutoFilterOperator.xlAnd
filterItem.Criteria1= '>=100'
filterItem.Criteria2= '<=200'
autoFilter.ApplyFilter()
```

6. 名前に「Zhang」を含む、または「Li」で始まる行にフィルター

```js
const autoFilter = ActiveSheet.AutoFilter
const filterItem = autoFilter.Filters.Item(2)
filterItem.Operator = Enum.XlAutoFilterOperator.xlOr
filterItem.Criteria1= '=*Zhang*'
filterItem.Criteria2= '=Li*'
autoFilter.ApplyFilter()
```

## Filters

フィルターオブジェクトのコレクションです。

シート上で設定できる各列のフィルターが含まれます。

#### データ型

Filters - フィルターのコレクション

### Count

フィルターオブジェクトの数です。

#### データ型

数値 - 現在のシート上のフィルター数

#### 例

```js
const autoFilter = Application.ActiveSheet.AutoFilter
const filters = autoFilter.Filters
const count = filters.Count
console.log(count)
```

### Each()

すべてのフィルターを反復し、コールバックを実行します。

#### パラメーター

| パラメーター | 型 | 既定値 | 必須 | 説明 |
| --- | --- | --- | --- | --- |
| callback | 関数 | null | はい | JS の配列 forEach に類似 |

#### 例

```js
const autoFilter = Application.ActiveSheet.AutoFilter
const filters = autoFilter.Filters
filters.Each(item => {
  console.log(item.Operator)
})
```

### Item()

インデックスにより、列のフィルターを取得します。

#### パラメーター

| パラメーター | 型 | 既定値 | 必須 | 説明 |
| --- | --- | --- | --- | --- |
| index | 文字列/数値 |  | はい | シート上の物理列インデックス。1 から始まる |

#### 戻り値

Filter - そのインデックスの Filter オブジェクト

#### 例

```js
const autoFilter = Application.ActiveSheet.AutoFilter
const filters = autoFilter.Filters
const filter1 = filters.Item(1)
```

### Item().Operator

フィルター演算子。[XlAutoFilterOperator](Enum.md#xlautofilteroperator) を参照してください。

#### データ型

[XlAutoFilterOperator](Enum.md#xlautofilteroperator) - XlAutoFilterOperator 列挙型

#### 例

```js
const autoFilter = Application.ActiveSheet.AutoFilter
const filter1 = autoFilter.Filters.Item(1)
filter1.Operator = Enum.XlAutoFilterOperator.xlFilterCellColor
```

### Item().Criteria1

第 1 条件です。空セルに一致させるには "="、空でないセルには "<>" を使います。省略するとすべての値に一致します。Operator が xlTop10Items の場合、Criteria1 は件数です。

#### データ型

Variant - フィルターの種類に依存

### Item().Criteria2

第 2 条件です。Criteria1 および Operator と組み合わせて使います。日付フィルター（日/月/年）にも使われます。フィルターを記述する配列の場合もあります: Array(Level, Date)。Level は 0〜2（年、月、日）、Date は範囲内の有効な日付です。

#### データ型

Variant - フィルターの種類に依存

## Range

AutoFilter が覆う領域です。

#### データ型

[Range](Range.md) - Range オブジェクト

#### 例

```js
const autoFilter = Application.ActiveSheet.AutoFilter
const range = autoFilter.Range
console.log(range.Count)
```

## ApplyFilter()

演算子と条件を設定したあと、AutoFilter を適用します。

#### 例

```js
const autoFilter = Application.ActiveSheet.AutoFilter
autoFilter.Filters.Item(1).Operator =
  Enum.XlAutoFilterOperator.xlFilterCellColor
autoFilter.Filters.Item(1).Criteria1 = '#44546A'
autoFilter.ApplyFilter()
```

## ShowAllData()

すべてのフィルターを解除し、すべての行を表示します。

#### 例

```js
const autoFilter = Application.ActiveSheet.AutoFilter
autoFilter.ShowAllData()
```
