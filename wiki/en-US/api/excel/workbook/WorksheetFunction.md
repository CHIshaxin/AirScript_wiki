---
title: "WorksheetFunction"
source: https://airsheet.wps.cn/docs/en-US/api/excel/workbook/WorksheetFunction.html
updated: 1789700916000
hash: Cexh2PSz
---

# WorksheetFunction

The worksheet function object is a built-in Excel object that exposes common Excel functions such as Sum, Average, Min, and Max. Use WorksheetFunction in script to call these functions for analysis and calculation.

You can assign results to variables or call functions inline for numeric work, text handling, and more complex analysis.

For example, Sum totals a column, Average returns the mean, Max returns the largest value, and Min returns the smallest.

#### Methods

| Method | Return type | Description |
|---|---|---|
| [Average()](#average) | Number | Average of numbers in a range |
| [AverageIf()](#averageif) | Number | Average of cells that meet a condition |
| [Small()](#small) | Number | k-th smallest value in a data set |
| [Large()](#large) | Number | k-th largest value in an array or column |
| [Min()](#min) | Number | Minimum value in an array or column |
| [Max()](#max) | Number | Maximum value in an array or column |
| [Sum()](#sum) | Number | Sum of numbers in a range |

## Average()

Average of numbers in the specified range.

### Parameters

| Property | Data type | Default | Required | Description |
|---|---|---|---|---|
| range | [Range](/docs/en-US/api/excel/workbook/Range.html) | null | Yes | Range to average |

### Return value

Number - Average of numbers in the range

```js
// Table with names in column A and scores in column B; average column B:

const range = Range('B:B')
const average = WorksheetFunction.Average(range)
console.log(average)
```

## AverageIf()

Average of cells that meet a condition.

### Parameters

| Property | Data type | Default | Required | Description |
|---|---|---|---|---|
| range | [Range](/docs/en-US/api/excel/workbook/Range.html) | null | Yes | Range to evaluate |
| criteria | String | null | Yes | Condition as number, expression, reference, or text (e.g. 32, "32", ">32", "apples", or B4) |
| average_range | [Range](/docs/en-US/api/excel/workbook/Range.html) | null | No | Actual cells to average; if omitted, range is used |

### Return value

Number - Average of matching numbers

```js
// Average scores in B where name in A is "Alice"

const rangeA = Range('A:A')
const rangeB = Range('B:B')
const average = WorksheetFunction.AverageIf(rangeA, 'Alice', rangeB)
console.log(average)
```

## Small()

Returns the k-th smallest value in a data set.

If the array is empty, Small() returns `#NUM!`.

If k <= 0 or k exceeds the number of data points, Small returns `#NUM!`.

If n is the number of points, SMALL(array,1) is the minimum and SMALL(array,n) is the maximum.

### Parameters

| Property | Data type | Default | Required | Description |
|---|---|---|---|---|
| range | [Range](/docs/en-US/api/excel/workbook/Range.html) | null | Yes | Array or column from which to take the k-th smallest |
| k | Number | null | Yes | Rank of the smallest value; k must be > 0 and <= number of elements |

### Return value

Number - The k-th smallest value

### Example

```js
// Third smallest score in column B

const range = Range('B:B')
const third = WorksheetFunction.Small(range, 3)
console.log(third)
```

## Large()

Returns the k-th largest value in an array or column (e.g. top score, second, third).

If the array is empty, Large returns `#NUM!`.

If k <= 0 or k exceeds the number of points, Large returns `#NUM!`.

If there are n points, LARGE(array,1) is the maximum and LARGE(array,n) is the minimum.

### Parameters

| Property | Data type | Default | Required | Description |
|---|---|---|---|---|
| range | [Range](/docs/en-US/api/excel/workbook/Range.html) | null | Yes | Array or column from which to take the k-th largest |
| k | Number | null | Yes | Rank of the largest value; k must be > 0 and <= number of elements |

### Return value

Number - The k-th largest value

### Example

```js
// Fourth largest score in column B

const range = Range('B:B')
const fourth = WorksheetFunction.Large(range, 4)
console.log(fourth)
```

## Min()

Minimum value in an array or column.

Arguments may be numbers, names, arrays, or references.

Logical values and text representations of numbers typed directly in the argument list are included.

If an argument is an array or reference, only numbers are used. Blank cells, logical values, and text in the array or reference are ignored.

If there are no numbers, Min returns 0.

Errors or text that cannot be converted to numbers produce an error.

### Parameters

| Property | Data type | Default | Required | Description |
|---|---|---|---|---|
| range | [Range](/docs/en-US/api/excel/workbook/Range.html) | null | Yes | Numbers from which to take the minimum |

### Return value

Number - Minimum value

### Example

```js
// Minimum score in column B

const range = Range('B:B')
const min = WorksheetFunction.Min(range)
console.log(min)
```

## Max()

Maximum value in an array or column.

Arguments may be numbers, names, arrays, or references.

Logical values and text representations of numbers typed directly in the argument list are included.

If an argument is an array or reference, only numbers are used. Blank cells, logical values, and text in the array or reference are ignored.

If there are no numbers, Max returns 0.

Errors or text that cannot be converted to numbers produce an error.

### Parameters

| Property | Data type | Default | Required | Description |
|---|---|---|---|---|
| range | [Range](/docs/en-US/api/excel/workbook/Range.html) | null | Yes | Array or column for the maximum |

### Return value

Number - Maximum value

### Example

```js
// Maximum score in column B

const range = Range('B:B')
const max = WorksheetFunction.Max(range)
console.log(max)
```

## Sum()

Sum of numbers in a range.

Numbers, logical values, and text representations of numbers typed directly in the argument list are included.

If an argument is an array or reference, only numbers in that array or reference are summed. Empty cells, logical values, and text are ignored.

Errors or text that cannot be converted to numbers produce an error.

### Parameters

| Property | Data type | Default | Required | Description |
|---|---|---|---|---|
| Array | [Range](/docs/en-US/api/excel/workbook/Range.html) | null | Yes | Arguments to sum |

### Return value

Number - Sum of all numeric values

### Example

```js
// Sum of scores in column B

const range = Range('B:B')
const sum = WorksheetFunction.Sum(range)
console.log(sum)
```

## Match()

Returns the relative position of a value in a range.

Use Match when you need the position of an item, not the item itself, instead of Lookup(Object, Object, Object).

Lookup_value is the value to find in lookup_array (for example, a name in a phone book when you ultimately need the phone number).

Lookup_value may be a number, text, logical value, or a reference to one of those.

If match_type is 1, Match finds the largest value less than or equal to lookup_value; lookup_array must be sorted ascending: ... -2, -1, 0, 1, 2, ..., A-Z, false, true.

If match_type is 0, Match finds the first value exactly equal to lookup_array may be in any order.

If match_type is -1, Match finds the smallest value greater than or equal to lookup_value; lookup_array must be sorted descending: true, false, Z-A, ... 2, 1, 0, -1, -2, ... If match_type is omitted, 1 is assumed.

Match returns the position in lookup_array, not the value. For example, MATCH("b",{"a","b","c"},0) returns 2.

Match is not case-sensitive for text.

If no match is found, Match returns #N/A.

If match_type is 0 and lookup_value is text, wildcards ? and * are allowed. ? matches one character; * matches any sequence. To find a literal ? or *, prefix with ~.

### Parameters

| Property | Data type | Default | Required | Description |
|---|---|---|---|---|
| Lookup_value | Object | null | Yes | Value to find |
| Lookup_array | [Range](/docs/en-US/api/excel/workbook/Range.html) | null | Yes | Contiguous range that may contain the value; must be an array or reference |
| Match_type | Object | null | No | -1, 0, or 1: how lookup_value is compared to lookup_array in Excel |

### Return value

Number - Position of lookup_value in lookup_array

### Example

```js
// Find score for "Alice" using names in column A

const range = Range('A:A')
const rowIndex = WorksheetFunction.Match('Alice', range, 0)
const score = Range(`B${rowIndex}`).Value
console.log(score)
```
