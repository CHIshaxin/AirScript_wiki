---
title: "Filter 对象"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Filter/Filter%20%E5%AF%B9%E8%B1%A1.html
---
# Filter (对象)

代表单个列的筛选。

## 说明

**[Filter](./../Filter/Filter 对象.html)** 对象是 **[Filters](./../Filters/Filters 对象.html)** 集合的成员。 **[Filters](./../Filters/Filters 对象.html)** 集合包含自动筛选区域中的所有筛选。

使用 **Filters**( *index*)（其中 *index* 是筛选的标题或索引号）可返回单个 **[Filter](./../Filter/Filter 对象.html)** 对象。下例将一个变量设置为工作表 Crew 上的筛选区域中第一列的筛选的 **On** 属性的值。

```javascript
function test()
{
    let w = Application.Worksheets.Item("Crew")
    if(w.AutoFilterMode) {
        filterIsOn = w.AutoFilter.Filters.Item(1).On
    }
}
```

注意： **[Filter](./../Filter/Filter 对象.html)** 对象的所有属性都是只读的。要设置这些属性，请手动应用自动筛选，或使用 **[Range](./../Range/Range 对象.html)** 对象的 **AutoFilter** 方法，如下例中所示。

```javascript
function test() {
    let w = Application.Worksheets.Item("Crew")
    w.Cells.AutoFilter(2, "Crucial", xlOr, "Important")
}
```
