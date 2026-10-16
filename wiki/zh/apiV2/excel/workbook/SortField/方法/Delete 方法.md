---
title: "Delete (方法)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/SortField/%E6%96%B9%E6%B3%95/Delete%20%E6%96%B9%E6%B3%95.html
updated: 1789700916000
hash: C1OEf9R2
---

# Delete (方法)

从 **[SortFields](./../../SortFields/SortFields 对象.html)** 集合中删除指定的 **[SortField](./../../SortField/SortField 对象.html)** 对象。

## 示例

```javascript
/*本示例在活动工作表创建两个新的排序字段，并删除活动工作表第二个SortField对象。*/
function test() {
    let sort = Application.Sheets.Item(1).Sort.SortFields
    sort.Add(Range("A1:A5"), xlSortOnValues, xlAscending)
    sort.Add(Range("B1:B5"), xlSortOnValues, xlAscending)
    sort.Item(2).Delete()
}
```

```javascript
/*本示例判断如果第一张工作表SortFields集合中的SortField对象数目大于1，则删除第一个SortField对象。*/
function test() {
    let sort = Application.Sheets.Item(1).Sort.SortFields
    if (sort.Count > 1) {
        sort.Item(1).Delete()
    }
}
```
