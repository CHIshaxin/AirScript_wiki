---
title: "SortField"
source: https://airsheet.wps.cn/docs/api/excel/workbook/SortField.html
updated: 1789700916000
hash: sAK9jPet
---

# SortField

代表与 **Sort** 对象关联的排序字段对象

#### 属性列表

| 属性名 | 数据类型 | 简介 |
|---|---|---|
| [CustomOrder](#customorder) | Variant | 指定对字段进行排序的自定义次序 |
| [DataOption](#dataoption) | [XlSortDataOption](/docs/api/excel/workbook/Enum.html#xlsortdataoption) | 指定如何在 SortField 对象中指定的区域中对文本进行排序 |
| [Key](#key) | [Range](/docs/api/excel/workbook/Range.html) | 指定排序字段，该字段确定要排序的值 |
| [Order](#order) | [XlSortOrder](/docs/api/excel/workbook/Enum.html#xlsortorder) | 确定关键字所指定的值的排序次序 |
| [Priority](#priority) | Number | 指定排序字段的优先级 |
| [SortOn](#sorton) | [XlSortOn](/docs/api/excel/workbook/Enum.html#xlsorton) | 返回或设置要排序的单元格的属性 |
| [SortOnValue](#sortonvalue) | Object | 返回针对指定的 SortField 对象执行排序的值 |

#### 方法列表

| 方法名 | 返回类型 | 简介 |
|---|---|---|
| [Delete()](#delete) | undefined | 从 SortFields 集合中删除指定的 SortField 对象 |
| [ModifyKey()](#modifykey) | undefined | 修改字段中按其排序的键值 |

## CustomOrder

指定对字段进行排序的自定义次序，可读/写

#### 数据类型

Variant - 自定义次序对象

## DataOption

指定如何在 SortField 对象中指定的区域中对文本进行排序，可读/写

#### 数据类型

[XlSortDataOption](/docs/api/excel/workbook/Enum.html#xlsortdataoption) - 文本排序方式枚举

## Key

指定排序字段，该字段确定要排序的值，此为只读属性

#### 数据类型

[Range](/docs/api/excel/workbook/Range.html) - 区域对象

## Order

确定关键字所指定的值的排序次序，可读/写

#### 数据类型

[XlSortOrder](/docs/api/excel/workbook/Enum.html#xlsortorder) - 排序次序枚举

## Priority

指定排序字段的优先级，可读/写

#### 数据类型

Number

## SortOn

返回或设置要排序的单元格的属性，可读/写

#### 数据类型

[XlSortOn](/docs/api/excel/workbook/Enum.html#xlsorton)

## SortOnValue

返回针对指定的 SortField 对象执行排序的值，此为只读属性

#### 数据类型

Object

## Delete()

从 SortFields 集合中删除指定的 SortField 对象

#### 返回值

undefined

## ModifyKey()

修改字段中按其排序的键值

#### 参数

| 属性 | 数据类型 | 默认值 | 必填 | 说明 |
|---|---|---|---|---|
| Key | [Range](/docs/api/excel/workbook/Range.html) | null | 是 | 指定要修改的键 |

#### 返回值

undefined
