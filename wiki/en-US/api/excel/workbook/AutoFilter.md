---
title: "AutoFilter"
source: https://airsheet.wps.cn/docs/en-US/api/excel/workbook/AutoFilter.html
updated: 1789700916000
hash: QPThMj-_
---

# AutoFilter

AutoFilter for a worksheet. Use it to filter data and find values quickly. You can combine filters on one or more columns.

Filtering lets you control what to show and what to hide. When values in one or more columns do not meet the criteria, the entire row is hidden. You can filter numbers or text, or filter by cell fill or font color.

Developers use the [Filters](#filters) property to get the collection of column filters. The [Range](#range) property returns a [Range](Range.md) for the whole filtered region.

Note

To create an AutoFilter on a sheet, turn on filtering in the toolbar manually, or use the [AutoFilter](Range.md#autofilter) method on a [Range](Range.md) to enable AutoFilter for a region.

See the lists below for the AutoFilter object's properties and methods.

#### Properties

| Property | Data type | Description |
| --- | --- | --- |
| [Filters](#filters) | [Filters](#filters) | Filter collection |
| [Range](#range) | [Range](Range.md) | Filtered range |

#### Methods

| Method | Return type | Description |
| --- | --- | --- |
| [ApplyFilter()](#applyfilter) | undefined | Apply filters to the current sheet |
| [ShowAllData()](#showalldata) | undefined | Clear all filters and show all rows |

#### Usage examples

Typical AutoFilter scenarios on a worksheet:

Assume we have the following sheet and want to filter to the rows we need.

1. Filter rows where the name is `Kim`

```js
// AutoFilter object
const autoFilter = ActiveSheet.AutoFilter
// Filter for name column
const filter2 = autoFilter.Filters.Item(2)
// Value filter
filter2.Operator = Enum.XlAutoFilterOperator.xlFilterValues
filter2.Criteria1 = ['Kim']
autoFilter.ApplyFilter()
```

2. Filter Chinese scores in the top 20%

```js
const autoFilter = ActiveSheet.AutoFilter
const filterItem = autoFilter.Filters.Item(5)
filterItem.Operator = Enum.XlAutoFilterOperator.xlTop10Percent
filterItem.Criteria1 = '20'
autoFilter.ApplyFilter()
```

3. Filter top 10 math scores

```js
const autoFilter = ActiveSheet.AutoFilter
const filterItem = autoFilter.Filters.Item(6)
filterItem.Operator = Enum.XlAutoFilterOperator.xlTop10Items
filterItem.Criteria1 = '10'
autoFilter.ApplyFilter()
```

4. Filter name column where cell fill is red

```js
const autoFilter = ActiveSheet.AutoFilter
const filterItem = autoFilter.Filters.Item(2)
filterItem.Operator = Enum.XlAutoFilterOperator.xlFilterCellColor
filterItem.Criteria1 = '#FF0000'
autoFilter.ApplyFilter()
```

5. Filter Chinese scores between 100 and 200

```js
const autoFilter = ActiveSheet.AutoFilter
const filterItem = autoFilter.Filters.Item(5)
filterItem.Operator = Enum.XlAutoFilterOperator.xlAnd
filterItem.Criteria1= '>=100'
filterItem.Criteria2= '<=200'
autoFilter.ApplyFilter()
```

6. Filter names that contain "Zhang" or start with "Li"

```js
const autoFilter = ActiveSheet.AutoFilter
const filterItem = autoFilter.Filters.Item(2)
filterItem.Operator = Enum.XlAutoFilterOperator.xlOr
filterItem.Criteria1= '=*Zhang*'
filterItem.Criteria2= '=Li*'
autoFilter.ApplyFilter()
```

## Filters

Collection of filter objects.

Contains every column filter you can configure on the sheet.

#### Data type

Filters - Filter collection

### Count

Number of filter objects.

#### Data type

Number - Number of filters on the current sheet

#### Example

```js
const autoFilter = Application.ActiveSheet.AutoFilter
const filters = autoFilter.Filters
const count = filters.Count
console.log(count)
```

### Each()

Iterates all filters and runs a callback.

#### Parameters

| Property | Data type | Default | Required | Description |
| --- | --- | --- | --- | --- |
| callback | Function | null | Yes | Similar to JS array forEach |

#### Example

```js
const autoFilter = Application.ActiveSheet.AutoFilter
const filters = autoFilter.Filters
filters.Each(item => {
  console.log(item.Operator)
})
```

### Item()

Gets the filter for a column by index.

#### Parameters

| Property | Data type | Default | Required | Description |
| --- | --- | --- | --- | --- |
| index | String/Number |  | Yes | Physical column index on the sheet, starting at 1 |

#### Return type

Filter - Filter object for that index

#### Example

```js
const autoFilter = Application.ActiveSheet.AutoFilter
const filters = autoFilter.Filters
const filter1 = filters.Item(1)
```

### Item().Operator

Filter operator; see [XlAutoFilterOperator](Enum.md#xlautofilteroperator).

#### Data type

[XlAutoFilterOperator](Enum.md#xlautofilteroperator) - XlAutoFilterOperator enum

#### Example

```js
const autoFilter = Application.ActiveSheet.AutoFilter
const filter1 = autoFilter.Filters.Item(1)
filter1.Operator = Enum.XlAutoFilterOperator.xlFilterCellColor
```

### Item().Criteria1

First criterion. Use "=" to match empty cells, or "<>" for non-empty. If omitted, all values match. If Operator is xlTop10Items, Criteria1 is the item count.

#### Data type

Variant - Depends on filter type

### Item().Criteria2

Second criterion, combined with Criteria1 and Operator. Also used for date filters (day/month/year). May be an array describing the filter: Array(Level, Date) where Level is 0-2 (year, month, day) and Date is a valid date in the range.

#### Data type

Variant - Depends on filter type

## Range

Region covered by AutoFilter.

#### Data type

[Range](Range.md) - Range object

#### Example

```js
const autoFilter = Application.ActiveSheet.AutoFilter
const range = autoFilter.Range
console.log(range.Count)
```

## ApplyFilter()

Applies the AutoFilter after you set operator and criteria.

#### Example

```js
const autoFilter = Application.ActiveSheet.AutoFilter
autoFilter.Filters.Item(1).Operator =
  Enum.XlAutoFilterOperator.xlFilterCellColor
autoFilter.Filters.Item(1).Criteria1 = '#44546A'
autoFilter.ApplyFilter()
```

## ShowAllData()

Clears all filters and shows all rows.

#### Example

```js
const autoFilter = Application.ActiveSheet.AutoFilter
autoFilter.ShowAllData()
```
