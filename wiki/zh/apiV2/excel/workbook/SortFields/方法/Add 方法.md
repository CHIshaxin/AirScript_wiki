---
title: "Add (方法)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/SortFields/%E6%96%B9%E6%B3%95/Add%20%E6%96%B9%E6%B3%95.html
updated: 1789700916000
hash: Dr4z9-qb
---

# Add (方法)

创建新的排序字段，并返回一个 **[SortFields](./../../SortFields/SortFields 对象.html)** 对象。

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| Key | Range | 必填 | 指定用于排序的键值。 |
| SortOn | any | 可选 | 要进行排序的字段。 |
| Order | any | 可选 | 指定排序次序。 |
| CustomOrder | any | 可选 | 指定是否应使用自定义排序次序。 |
| DataOption | any | 可选 | 指定数据选项。 |

## 返回值

SortField

## 示例

```javascript
/*本示例在活动工作表创建新的排序字段，并显示排序字段的键值地址。*/
function test() {
    let sort = Application.ActiveSheet.Sort.SortFields.Add(Range("A1:D1"), xlSortOnCellColor)
    console.log(sort.Key.Address())
}
```

```javascript
/*本示例在第一张工作表创建新的排序字段*/
function test() {
    Application.Sheets.Item(1).Sort.SortFields.Add(Range("A1:A11"), xlSortOnValues, xlAscending, null, xlSortNormal)
}
```
