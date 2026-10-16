---
title: "SortField"
source: https://airsheet.wps.cn/docs/ja-JP/api/excel/workbook/SortField.html
updated: 1789700916000
hash: CSCKW6Iv
---

# SortField

**Sort** オブジェクトに関連付けられた並べ替えフィールドを表します。

#### プロパティ

| プロパティ | 型 | 説明 |
|---|---|---|
| [CustomOrder](#customorder) | Variant | フィールドのカスタム並べ替え順 |
| [DataOption](#dataoption) | [XlSortDataOption](/docs/ja-JP/api/excel/workbook/Enum.html#xlsortdataoption) | SortField で指定された範囲内のテキストの並べ替え方法 |
| [Key](#key) | [Range](/docs/ja-JP/api/excel/workbook/Range.html) | 並べ替え対象の値を決定する並べ替えキーの範囲 |
| [Order](#order) | [XlSortOrder](/docs/ja-JP/api/excel/workbook/Enum.html#xlsortorder) | キーの並べ替え順 |
| [Priority](#priority) | Number | この並べ替えフィールドの優先順位 |
| [SortOn](#sorton) | [XlSortOn](/docs/ja-JP/api/excel/workbook/Enum.html#xlsorton) | 並べ替えに使用するセルのプロパティ |
| [SortOnValue](#sortonvalue) | Object | この SortField の並べ替えに使用される値 |

#### メソッド

| メソッド | 戻り値の型 | 説明 |
|---|---|---|
| [Delete()](#delete) | undefined | この SortField を SortFields コレクションから削除します |
| [ModifyKey()](#modifykey) | undefined | 並べ替えに使用するキー範囲を変更します |

## CustomOrder

フィールドのカスタム並べ替え順です。読み書き可能です。

#### データ型

Variant - カスタム順序オブジェクト

## DataOption

SortField で指定された範囲内のテキストの並べ替え方法です。読み書き可能です。

#### データ型

[XlSortDataOption](/docs/ja-JP/api/excel/workbook/Enum.html#xlsortdataoption) - テキストの並べ替えオプション

## Key

並べ替え対象の値を決定する並べ替えキーの範囲です。読み取り専用です。

#### データ型

[Range](/docs/ja-JP/api/excel/workbook/Range.html) - Range オブジェクト

## Order

キー値の並べ替え順です。読み書き可能です。

#### データ型

[XlSortOrder](/docs/ja-JP/api/excel/workbook/Enum.html#xlsortorder) - 並べ替え順の列挙型

## Priority

この並べ替えフィールドの優先順位です。読み書き可能です。

#### データ型

Number

## SortOn

並べ替えに使用するセルのプロパティです。読み書き可能です。

#### データ型

[XlSortOn](/docs/ja-JP/api/excel/workbook/Enum.html#xlsorton)

## SortOnValue

この SortField の並べ替えに使用される値です。読み取り専用です。

#### データ型

Object

## Delete()

この SortField を SortFields コレクションから削除します。

#### 戻り値

undefined

## ModifyKey()

並べ替えに使用するキー範囲を変更します。

#### パラメーター

| パラメーター | 型 | 既定値 | 必須 | 説明 |
|---|---|---|---|---|
| Key | [Range](/docs/ja-JP/api/excel/workbook/Range.html) | null | はい | 新しいキー範囲 |

#### 戻り値

undefined
