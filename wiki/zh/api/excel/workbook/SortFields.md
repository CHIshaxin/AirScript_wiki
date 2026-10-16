---
title: "SortFields.md"
source: https://airsheet.wps.cn/docs/api/excel/workbook/SortFields.html
updated: 1789700916000
hash: BUuCwINh
---

### Count

返回集合中对象的数目，只读。

#### 数据类型

Number - 对象数目

### Add()

创建新的排序字段，并返回一个 SortFields 对象

#### 参数

| 属性 | 数据类型 | 默认值 | 必填 | 说明 |
| --- | --- | --- | --- | --- |
| Key | [Range](/docs/api/excel/workbook/Range.html) | null | 是 | 指定用于排序的键值 |
| SortOn | [XlSortOn](/docs/api/excel/workbook/Enum.html#xlsorton) | null | 否 | 要进行排序的字段 |
| Order | [XlSortOrder](/docs/api/excel/workbook/Enum.html#xlsortorder) | null | 否 | 指定排序次序 |
| CustomOrder | Variant | null | 否 | 指定是否应使用自定义排序次序 |
| DataOption | [XlSortDataOption](/docs/api/excel/workbook/Enum.html#xlsortdataoption) | null | 否 | 指定数据选项 |

#### 返回值

[SortField](/docs/api/excel/workbook/SortField.html) - 代表与 **Sort** 对象关联的排序字段对象

### Clear()

清除所有 SortFields 对象，在开始添加排序字段时，最好先调用一次此方法

#### 返回值

undefined

### Item()

返回一个 [SortField](/docs/api/excel/workbook/SortField.html) 对象，此为只读属性

#### 参数

| 属性 | 数据类型 | 默认值 | 必填 | 说明 |
| --- | --- | --- | --- | --- |
| Index | Number / String | null | 是 | 指定用于索引，默认从 1 开始 |

#### 返回值

[SortField](/docs/api/excel/workbook/SortField.html) - 代表与 **Sort** 对象关联的排序字段对象
