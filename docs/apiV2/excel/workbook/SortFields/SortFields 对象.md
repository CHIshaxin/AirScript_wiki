---
title: "SortFields 对象"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/SortFields/SortFields%20%E5%AF%B9%E8%B1%A1.html
---
# SortFields (对象)

**[SortFields](./../SortFields/SortFields 对象.html)** 集合是 **[SortField](./../SortField/SortField 对象.html)** 对象的集合。开发人员可以使用该集合存储工作簿、列表和自动筛选的排序状态。

## 示例

```javascript
/*本示例为活动工作表创建两个新的排序字段。*/
function test() {
    Application.ActiveSheet.Sort.SortFields.Add(Range("A1"), null, xlDescending)
    Application.ActiveSheet.Sort.SortFields.Add(Range("B1"), null, xlDescending)
}
```

```javascript
/*本示例清除第一张工作表上所有的 SortField 对象，并创建新的排序字段。*/
function test() {
    Application.Sheets.Item(1).Sort.SortFields.Clear()
    Application.Sheets.Item(1).Sort.SortFields.Add(Range("A1:A11"), xlSortOnValues, xlAscending)
}
```
