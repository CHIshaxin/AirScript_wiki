---
title: "Sort 方法"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Range/%E6%96%B9%E6%B3%95/Sort%20%E6%96%B9%E6%B3%95.html
---
# Sort (方法)

对值区域进行排序。

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| Key1 | any | 可选 | 指定第一排序字段，作为区域名称（字符串）或 **[Range](./../../Range/Range 对象.html)** 对象；确定要排序的值。 |
| Order1 | Default<XlSortOrder, 1> | 可选 | 确定 Key1 中指定的值的排序次序。 |
| Key2 | any | 可选 | 第二排序字段；对数据透视表进行排序时不能使用。 |
| Type | any | 可选 | 指定要排序的元素。 |
| Order2 | Default<XlSortOrder, 1> | 可选 | 确定 Key2 中指定的值的排序次序。 |
| Key3 | any | 可选 | 第三排序字段；对数据透视表进行排序时不能使用。 |
| Order3 | Default<XlSortOrder, 1> | 可选 | 确定 Key3 中指定的值的排序次序。 |
| Header | Default<XlYesNoGuess, 2> | 可选 | 指定第一行是否包含标题信息。**xlNo** 是默认值；如果希望由 ET 尝试确定标题，则指定 **xlGuess**。 |
| OrderCustom | any | 可选 | 指定在自定义排序次序列表中的基于一的整数偏移。 |
| MatchCase | any | 可选 | 设置为 **True**，则执行区分大小写的排序，设置为 **False**，则执行不区分大小写的排序；不能用于数据透视表。 |
| Orientation | Default<XlSortOrientation, 2> | 可选 | 指定是应按行（默认）还是按列进行排序。 将 xlSortColumns 值设置为 1 以按列排序。 将 xlSortRows 值设置为 2 以按行排序， (这是默认值) 。 |
| SortMethod | Default<XlSortMethod, 1> | 可选 | 指定排序方法。 |
| DataOption1 | Default<XlSortDataOption, 0> | 可选 | 指定 Key1 中所指定区域中的文本的排序方式；不应用于数据透视表排序。 |
| DataOption2 | Default<XlSortDataOption, 0> | 可选 | 指定 Key2 中所指定区域中的文本的排序方式；不应用于数据透视表排序。 |
| DataOption3 | Default<XlSortDataOption, 0> | 可选 | 指定 Key3 中所指定区域中的文本的排序方式；不应用于数据透视表排序。 |

## 返回值

Variant

## 示例

```javascript
/*本示例对 A1:B10 单元格区域根据 A1 所在列进行降序排序（区分大小写）*/
function test() {
    Range("A1:B10").Sort(Range("A1"), xlDescending, null, null, null, null, null, null, null, true)
}
```

```javascript
/*本示例以 A1 所在列为第一字段，降序；B1 所在列为第二字段，升序；自动判断是否包含标题，进行排序*/
function test() {
    Range("A1:B10").Sort(Range("A1"), xlDescending, Range("B1"), null, xlAscending, null, null, xlGuess)
}
```
