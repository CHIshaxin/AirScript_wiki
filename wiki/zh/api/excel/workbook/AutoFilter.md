---
title: "AutoFilter"
source: https://airsheet.wps.cn/docs/api/excel/workbook/AutoFilter.html
updated: 1789700916000
hash: 7isl1q-w
---

# AutoFilter

代表对指定工作表的自动筛选，可通过此对象对工作表中的数据进行筛选，快速找到想要的值。可以组合一列或多列数据进行筛选。

使用筛选，不仅可以控制想要查看的内容，还可以控制想要排除的内容。在进行数据筛选时，如果一列或多列中的数值不能满足筛选条件，整行数据都会隐藏起来。可以对数值或文本值进行筛选，也可以对背景或文本应用颜色格式的单元格按颜色进行筛选。

开发者可以通过[Filters](#filters) 属性获取由各个列筛选组成的集合。 使用[Range](#range)属性可返回代表整个筛选区域的[Range](/docs/api/excel/workbook/Range.html)对象。

注意

若要为工作表创建 AutoFilter 对象，必须在工具栏手动开启筛选功能或者使用[Range](/docs/api/excel/workbook/Range.html)对象的[AutoFilter](/docs/api/excel/workbook/Range.html#autofilter)方法为工作表上的某个区域启用自动筛选。

AutoFilter 对象的具体属性和方法请参阅下方的列表。

#### 属性列表

| 属性名 | 数据类型 | 简介 |
|---|---|---|
| [Filters](#filters) | [Filters](#filters) | 筛选对象集合 |
| [Range](#range) | [Range](/docs/api/excel/workbook/Range.html) | 筛选区域 |

#### 方法列表

| 方法名 | 返回类型 | 简介 |
|---|---|---|
| [ApplyFilter()](#applyfilter) | undefined | 应用筛选到当前工作表 |
| [ShowAllData()](#showalldata) | undefined | 清除所有筛选条件，显示所有数据 |

#### 应用示例

以下为您展示自动筛选对象的在工作表内的一些常见应用场景：

假设我们有如下的工作表，现在需要对它进行筛选操作，以显示我们期望的数据。

1. 筛选出姓名为`金小獴`的数据

```js
// 获取自动筛选对象
const autoFilter = ActiveSheet.AutoFilter
// 获取姓名列的筛选对象
const filter2 = autoFilter.Filters.Item(2)
// 设置筛选类型为值筛选
filter2.Operator = Enum.XlAutoFilterOperator.xlFilterValues
// 设置筛选的值
filter2.Criteria1 = ['金小獴']
// 应用筛选
autoFilter.ApplyFilter()
```

2. 筛选出语文成绩在前 20%的数据

```js
// 获取自动筛选对象
const autoFilter = ActiveSheet.AutoFilter
// 获取语文列的筛选对象
const filterItem = autoFilter.Filters.Item(5)
// 设置筛选类型为头部百分比筛选
filterItem.Operator = Enum.XlAutoFilterOperator.xlTop10Percent
// 设置筛选条件为20%
filterItem.Criteria1 = '20'
// 应用筛选
autoFilter.ApplyFilter()
```

3. 筛选出数学前十名的数据

```js
// 获取自动筛选对象
const autoFilter = ActiveSheet.AutoFilter
// 获取数学列的筛选对象
const filterItem = autoFilter.Filters.Item(6)
// 设置筛选类型为头部筛选
filterItem.Operator = Enum.XlAutoFilterOperator.xlTop10Items
// 设置筛选条件为10，即前十名
filterItem.Criteria1 = '10'
// 应用筛选
autoFilter.ApplyFilter()
```

4. 将姓名列中单元格为红色的筛选出来

```js
// 获取自动筛选对象
const autoFilter = ActiveSheet.AutoFilter
// 获取姓名列的筛选对象
const filterItem = autoFilter.Filters.Item(2)
// 设置筛选类型为单元格颜色筛选
filterItem.Operator = Enum.XlAutoFilterOperator.xlFilterCellColor
// 设置筛选条件为红色
filterItem.Criteria1 = '#FF0000'
// 应用筛选
autoFilter.ApplyFilter()
```

5. 筛选出语文分介于100至200之间的数据

```js
// 获取自动筛选对象
const autoFilter = ActiveSheet.AutoFilter
// 获取语文列的筛选对象
const filterItem = autoFilter.Filters.Item(5)
// 设置筛选类型为条件与筛选
filterItem.Operator = Enum.XlAutoFilterOperator.xlAnd
// 设置筛选条件
filterItem.Criteria1= '>=100'
filterItem.Criteria2= '<=200'
// 应用筛选
autoFilter.ApplyFilter()
```

6. 筛选第姓名包含'张三'或者开头等于'李四'的数据

```js
// 获取自动筛选对象
const autoFilter = ActiveSheet.AutoFilter
// 获取姓名列的筛选对象
const filterItem = autoFilter.Filters.Item(2)
// 设置筛选类型为条件或筛选
filterItem.Operator = Enum.XlAutoFilterOperator.xlOr
// 设置筛选条件
filterItem.Criteria1= '=*张三*'
filterItem.Criteria2= '=李四*'
// 应用筛选
autoFilter.ApplyFilter()
```

## Filters

筛选对象集合

该集合内包含所有可供操作的数据列筛选对象

#### 数据类型

Filters - Filter 集合

### Count

当前所有筛选器对象的数量

#### 数据类型

Number - 对应当前工作表所有筛选对象的数量

#### 示例

```js
// 获取自动筛选对象
const autoFilter = Application.ActiveSheet.AutoFilter
// 获取筛选对象集合
const filters = autoFilter.Filters
// 获取当前工作表所有筛选对象的数量
const count = filters.Count
console.log(count)
```

### Each()

遍历所有 Filters 并执行回调函数

#### 参数

| 属性 | 数据类型 | 默认值 | 必填 | 说明 |
|---|---|---|---|---|
| callback | Function | null | 是 | 类似 JS 数组的 forEach |

#### 示例

```js
// 获取自动筛选对象
const autoFilter = Application.ActiveSheet.AutoFilter
// 获取筛选对象集合
const filters = autoFilter.Filters
// 遍历筛选集合，执行回调函数
filters.Each(item => {
  console.log(item.Operator)
})
```

### Item()

根据索引选择对应的筛选对象

#### 参数

| 属性 | 数据类型 | 默认值 | 必填 | 说明 |
|---|---|---|---|---|
| index | String/Number |  | 是 | 对应工作表内的实际列号，索引从 1 开始 |

#### 返回类型

Filter - 返回对应索引的筛选对象

#### 示例

```js
// 获取自动筛选对象
const autoFilter = Application.ActiveSheet.AutoFilter
// 获取筛选对象集合
const filters = autoFilter.Filters
// 获取第一个筛选对象
const filter1 = filters.Item(1)
```

### Item().Operator

指定筛选类型，可使用的筛选类型请参照枚举值[XlAutoFilterOperator](/docs/api/excel/workbook/Enum.html#xlautofilteroperator)

#### 数据类型

[XlAutoFilterOperator](/docs/api/excel/workbook/Enum.html#xlautofilteroperator) - 枚举值 XlAutoFilterOperator

#### 示例

```js
// 获取自动筛选对象
const autoFilter = Application.ActiveSheet.AutoFilter
// 获取第一个筛选对象
const filter1 = autoFilter.Filters.Item(1)
// 设置筛选类型为单元格颜色
filter1.Operator = Enum.XlAutoFilterOperator.xlFilterCellColor
```

### Item().Criteria1

指定判断条件，使用“=”查找空字段，或者使用“<>”查找非空字段。如果忽略该参数，那么判断是全部。如果参数 Operator 是 xlTop10Items，那么参数 Criterial1 指定项目的数量

#### 数据类型

Variant - 根据指定的筛选类型确定

### Item().Criteria2

第二个判断条件。与 Criteria1 和 Operator 一起组合成复合筛选条件。 也用作日期字段的单一条件（按日、月或年筛选）。 后跟一个数组，该数组用于描述筛选 Array(Level, Date)。 其中，Level 为 0-2（年、月、日），Date 为筛选期内的一个有效日期

#### 数据类型

Variant - 根据指定的筛选类型确定

## Range

自动筛选的区域范围

#### 数据类型

[Range](/docs/api/excel/workbook/Range.html) - 区域对象

#### 示例

```js
// 获取自动筛选对象
const autoFilter = Application.ActiveSheet.AutoFilter
// 获取自动筛选的区域
const range = autoFilter.Range
// 打印该区域内的单元格数量
console.log(range.Count)
```

## ApplyFilter()

将自动筛选器应用于区域，在设置好筛选类型和筛选条件后，调用此方法来应用筛选

#### 示例

```js
// 获取自动筛选对象
const autoFilter = Application.ActiveSheet.AutoFilter
// 设置第一列的筛选类型为单元格颜色
autoFilter.Filters.Item(1).Operator =
  Enum.XlAutoFilterOperator.xlFilterCellColor
// 设置第一列的筛选条件是#44546A
autoFilter.Filters.Item(1).Criteria1 = '#44546A'
// 应用筛选到区域
autoFilter.ApplyFilter()
```

## ShowAllData()

清除所有筛选条件，显示所有数据

#### 示例

```js
// 获取自动筛选对象
const autoFilter = Application.ActiveSheet.AutoFilter

// 清除所有筛选条件
autoFilter.ShowAllData()
```
