---
title: "WorksheetFunction"
source: https://airsheet.wps.cn/docs/api/excel/workbook/WorksheetFunction.html
updated: 1789700916000
hash: CnHGKdCO
---

# WorksheetFunction

工作表函数对象是 Excel 中的一个内置对象，它包含了许多常用的 Excel 函数，例如 Sum、Average、Min、Max 等。使用 WorksheetFunction 对象可以在 VBA 中调用这些 Excel 函数，以实现对工作表数据的处理和分析。

通过 WorksheetFunction 对象，可以调用 Excel 函数并将其结果赋值给变量，也可以直接在代码中使用这些函数来进行数值计算、字符串处理等操作，可以实现复杂的数据处理和分析需求。

例如，可以使用 WorksheetFunction 对象的 SUM 函数来计算一列数字的总和，也可以使用 Average 函数来计算这列数字的平均值。使用 WorksheetFunction 对象的 Max 函数可以找到一列数字中的最大值，Min 函数可以找到一列数字中的最小值。

#### 方法列表

| 方法名 | 返回类型 | 简介 |
|---|---|---|
| [Average()](#average) | Number | 用于计算指定区域内数字的平均值 |
| [AverageIf()](#averageif) | Number | 用于计算指定区域内满足给定条件的所有单元格的平均值 |
| [Small()](#small) | Number | 返回数据集中第 k 个最小值 |
| [Large()](#large) | Number | 用于在一个数组或一列数据中返回第 k 个最大值 |
| [Min()](#min) | Number | 用于在一个数组或一列数据中返回最小值 |
| [Max()](#max) | Number | 用于在一个数组或一列数据中返回最大值 |
| [Sum()](#sum) | Number | 对某单元格区域中的所有数字求和 |

## Average()

用于计算指定区域内数字的平均值

### 参数

| 属性 | 数据类型 | 默认值 | 必填 | 说明 |
|---|---|---|---|---|
| range | [Range](/docs/api/excel/workbook/Range.html) | null | 是 | 指定要进行计算的区域 |

### 返回值

Number - 函数返回指定区域内数字的平均值

```js
// 假设有一个数据表，其中A列为学生姓名，B列为学生成绩。现在需要计算B列的平均分数，可以使用以下代码：

// 取出B列
const range = Range('B:B')
// 获取平均值
const average = WorksheetFunction.Average(range)
console.log(average)
```

## AverageIf()

用于计算指定区域内满足给定条件的所有单元格的平均值

### 参数

| 属性 | 数据类型 | 默认值 | 必填 | 说明 |
|---|---|---|---|---|
| range | [Range](/docs/api/excel/workbook/Range.html) | null | 是 | 要求其平均值的一个或多个单元格 |
| criteria | String | null | 是 | 定义将对哪些单元格求平均值的条件，其形式可以为数字、表达式、单元格引用或文本。 例如，条件可以表示为 32、“32”、“>32”、“apples”或 B4 |
| average_range | [Range](/docs/api/excel/workbook/Range.html) | null | 否 | 要求其平均值的实际单元格集合，如果省略，则使用 range |

### 返回值

Number - 函数返回指定区域内满足条件的数字的平均值

```js
// 假设有一个数据表，其中A列为学生姓名，B列为学生成绩。现在需要计算B列中A列为“张三”的平均分数，可以使用以下代码：

// 获取A列
const rangeA = Range('A:A')
// 获取B列
const rangeB = Range('B:B')
// 获取平均值
const average = WorksheetFunction.AverageIf(rangeA, '张三', rangeB)
console.log(average)
```

## Small()

返回数据集中第 k 个最小值。 使用此函数可以返回数据集中特定位置上的数值。

如果数组为空， Small() 将返回`#NUM！` 。

如果 k ≤ 0 或 k 超过数据点数， Small 将返回`#NUM！`。

如果 n 为数组中数据点的个数，则 SMALL(array,1) 等于最小值，SMALL(array,n) 等于最大值。

### 参数

| 属性 | 数据类型 | 默认值 | 必填 | 说明 |
|---|---|---|---|---|
| range | [Range](/docs/api/excel/workbook/Range.html) | null | 是 | 要从中获取第 k 个最小值的数组或数据列 |
| k | Number | null | 是 | 要返回的第 k 个最小值的位置。k 必须大于 0，小于等于数组或数据列中的元素个数 |

### 返回值

Number - 函数返回第 k 个最小值

### 示例

```js
// 假设有一个数据表，其中A列为学生姓名，B列为学生成绩。现在需要从B列中获取第3个最小值，可以使用以下代码：

// 取出B列
const range = Range('B:B')
// 获取第三小的值
const third = WorksheetFunction.Small(range, 3)
console.log(third)
```

## Large()

用于在一个数组或一列数据中返回第 k 个最大值。例如，可以使用 Large 返回最高、亚军或第三名的分数。

如果数组为空， 则 Large 返回`#NUM！`。

如果 k ≤ 0 或 k 大于数据点数， 则 Large 返回`#NUM！`。

如果区域中数据点的个数为 n，则函数 LARGE(array,1) 返回最大值，函数 LARGE(array,n) 返回最小值。

### 参数

| 属性 | 数据类型 | 默认值 | 必填 | 说明 |
|---|---|---|---|---|
| range | [Range](/docs/api/excel/workbook/Range.html) | null | 是 | 要从中获取第 k 个最大值的数组或数据列 |
| k | Number | null | 是 | 要返回的第 k 个最大值的位置。k 必须大于 0，小于等于数组或数据列中的元素个数 |

### 返回值

Number - 函数返回第 k 个最大值

### 示例

```js
// 假设有一个数据表，其中A列为学生姓名，B列为学生成绩。现在需要从B列中获取第4个最大值，可以使用以下代码：

// 取出B列
const range = Range('B:B')
// 获取第四大的值
const fourth = WorksheetFunction.Large(range, 4)
console.log(fourth)
```

## Min()

用于在一个数组或一列数据中返回最小值

参数可以是数字，也可以是包含数字的名称、数组或引用。

直接键入参数列表的数字的逻辑值和文本表示也包括在内。

如果参数为数组或引用，则只使用其中的数值。 数组或引用中的空白单元格、逻辑值或文本将被忽略。

如果参数不包含数字， 则 Min 返回 0。

如果参数为错误值或不能转换为数字的文本，则将导致错误。

### 参数

| 属性 | 数据类型 | 默认值 | 必填 | 说明 |
|---|---|---|---|---|
| range | [Range](/docs/api/excel/workbook/Range.html) | null | 是 | 要从中查找最小值的 n 个数字 |

### 返回值

Number - 函数返回数组或数据列中的最小值

### 示例

```js
// 假设有一个数据表，其中A列为学生姓名，B列为学生成绩。现在需要从B列中获取最小值，可以使用以下代码：

// 取出B列
const range = Range('B:B')
// 获取最小值
const min = WorksheetFunction.Min(range)
console.log(min)
```

## Max()

用于在一个数组或一列数据中返回最大值

参数可以是数字，也可以是包含数字的名称、数组或引用。

直接键入参数列表的数字的逻辑值和文本表示也包括在内。

如果参数为数组或引用，则只使用其中的数值。 数组或引用中的空白单元格、逻辑值或文本将被忽略。

如果参数不包含数字， 则 Max 返回 0 (零) 。

如果参数为错误值或不能转换为数字的文本，则将导致错误。

### 参数

| 属性 | 数据类型 | 默认值 | 必填 | 说明 |
|---|---|---|---|---|
| range | [Range](/docs/api/excel/workbook/Range.html) | null | 是 | 要从中获取最大值的数组或数据列 |

### 返回值

Number - 最大值

### 示例

```js
// 假设有一个数据表，其中A列为学生姓名，B列为学生成绩。现在需要从B列中获取最大值，可以使用以下代码：

// 取出B列
const range = Range('B:B')
// 获取最大值
const max = WorksheetFunction.Max(range)
console.log(max)
```

## Sum()

对某单元格区域中的所有数字求和

直接键入参数列表的数字、逻辑值和数字的文本表示也包括在内。

如果参数为数组或引用，则只有该数组或引用中的数字将被计算在内。 数组或引用中的空单元格、逻辑值或文本将被忽略。

如果参数为错误值或不能转换为数字的文本，则将导致错误。

### 参数

| 属性 | 数据类型 | 默认值 | 必填 | 说明 |
|---|---|---|---|---|
| Array | [Range](/docs/api/excel/workbook/Range.html) | null | 是 | 要对其求和的 n 个参数 |

### 返回值

Number - 函数返回数组或数据列中所有数值的总和

### 示例

```js
// 假设有一个数据表，其中A列为学生姓名，B列为学生成绩。现在需要获取B列学生成绩的和，可以使用以下代码：

// 取出B列
const range = Range('B:B')
// 获取成绩的和
const sum = WorksheetFunction.Sum(range)
console.log(sum)
```

## Match()

用于在一个区域中查找某个值，并返回该值在区域中的位置。

如果需要项在某个范围中的位置而不是项本身，请使用 Match 而不是 Lookup(Object, Object, Object) 函数。

Lookup_value 为需要在 Look_array 中查找的数值。 例如，如果要在电话簿中查找某人的电话号码，则应该将姓名作为查找值，但实际上需要的是电话号码。

Lookup_value 可以为值（数字、文本或逻辑值）或对数字、文本或逻辑值的单元格引用。

如果 match_type 为 1， Match 将查找小于或等于 lookup_value 的最大值。 Lookup_array 必须按升序排列：...-2、-1、0、1、2、...、A-Z、 false、 true。

如果 match_type 为 0， Match 将查找与 lookup_value 完全相等的第一个值。 Lookup_array 可以按任何顺序排列。

如果 match_type 为 -1， Match 将查找大于或等于 lookup_value 的最小值。 Lookup_array 必须按降序排列： true、 false、Z-A、...2、1、0、-1、-2、...等。如果省略 match_type，则假定为 1。

Match 返回匹配值在 lookup_array 中的位置，而不是值本身。 例如，MATCH("b",{"a","b","c"},0) 返回 2，即“b”在数组 {"a","b","c"} 中的相应位置。

匹配 文本值时，Match 不区分大写字母和小写字母。

如果 Match 未能找到匹配项，则返回#N/A 错误值。

如果 match_type 为 0 且 lookup_value 为文本，则可以在 lookup_value 中使用通配符、问号 (?) 和星号 (*)。 问号匹配任意单个字符；星号匹配任意字符序列。 如果要查找实际的问号或星号，则请在该字符前键入一个波形符 (~)。

### 参数

| 属性 | 数据类型 | 默认值 | 必填 | 说明 |
|---|---|---|---|---|
| Lookup_value | Object | null | 是 | 需要在表中查找的值 |
| Lookup_array | [Range](/docs/api/excel/workbook/Range.html) | null | 是 | 可能包含所要查找的值的连续单元格区域。 Lookup_array 必须为数组或数组引用 |
| Match_type | Object | null | 否 | 数字 -1、0 或 1。 Match_type 指明 Microsoft Excel 如何将 lookup_value 与 lookup_array 中的值进行匹配 |

### 返回值

Number - 函数返回 lookup_value 在 lookup_array 中的位置

### 示例

```js
// 假设有一个数据表，其中A列为学生姓名，B列为学生成绩。现在需要查找名为“张三”的学生成绩

// 取出A列
const range = Range('A:A')
// 查找名为“张三”的学生成绩的索引
const rowIndex = WorksheetFunction.Match('张三', range, 0)
// 获取“张三”的成绩
const score = Range(`B${rowIndex}`).Value
console.log(score)
```
