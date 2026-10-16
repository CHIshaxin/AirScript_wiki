---
title: "SortField"
source: https://airsheet.wps.cn/docs/en-US/api/excel/workbook/SortField.html
updated: 1789700916000
hash: C7CYXkmm
---

# SortField

Represents a sort field associated with a **Sort** object.

#### Properties

| Property | Data type | Description |
|---|---|---|
| [CustomOrder](#customorder) | Variant | Custom sort order for the field |
| [DataOption](#dataoption) | [XlSortDataOption](/docs/en-US/api/excel/workbook/Enum.html#xlsortdataoption) | How to sort text in the range specified by the SortField |
| [Key](#key) | [Range](/docs/en-US/api/excel/workbook/Range.html) | Sort key range that determines values to sort |
| [Order](#order) | [XlSortOrder](/docs/en-US/api/excel/workbook/Enum.html#xlsortorder) | Sort order for the key |
| [Priority](#priority) | Number | Priority of this sort field |
| [SortOn](#sorton) | [XlSortOn](/docs/en-US/api/excel/workbook/Enum.html#xlsorton) | Cell property to sort by |
| [SortOnValue](#sortonvalue) | Object | Value used to sort for this SortField |

#### Methods

| Method | Return type | Description |
|---|---|---|
| [Delete()](#delete) | undefined | Removes this SortField from the SortFields collection |
| [ModifyKey()](#modifykey) | undefined | Changes the key range used for sorting |

## CustomOrder

Custom sort order for the field. Read/write.

#### Data type

Variant - Custom order object

## DataOption

How to sort text in the range specified by the SortField. Read/write.

#### Data type

[XlSortDataOption](/docs/en-US/api/excel/workbook/Enum.html#xlsortdataoption) - Text sort option

## Key

Sort key range that determines values to sort. Read-only.

#### Data type

[Range](/docs/en-US/api/excel/workbook/Range.html) - Range object

## Order

Sort order for the key values. Read/write.

#### Data type

[XlSortOrder](/docs/en-US/api/excel/workbook/Enum.html#xlsortorder) - Sort order enum

## Priority

Priority of this sort field. Read/write.

#### Data type

Number

## SortOn

Cell property to sort by. Read/write.

#### Data type

[XlSortOn](/docs/en-US/api/excel/workbook/Enum.html#xlsorton)

## SortOnValue

Value used for sorting this SortField. Read-only.

#### Data type

Object

## Delete()

Removes this SortField from the SortFields collection.

#### Return value

undefined

## ModifyKey()

Changes the key range used for sorting.

#### Parameters

| Property | Data type | Default | Required | Description |
|---|---|---|---|---|
| Key | [Range](/docs/en-US/api/excel/workbook/Range.html) | null | Yes | New key range |

#### Return value

undefined
