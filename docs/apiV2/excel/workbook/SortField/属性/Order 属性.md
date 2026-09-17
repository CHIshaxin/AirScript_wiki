---
title: "Order 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/SortField/%E5%B1%9E%E6%80%A7/Order%20%E5%B1%9E%E6%80%A7.html
---
# Order (属性)

确定关键字所指定的值的排序次序。可读/写。

## 示例

```javascript
/*本示例在活动工作表创建新的排序字段，并设置该排序字段的排序次序为降序。*/
function test() {
    let sortfields = Application.ActiveSheet.Sort.SortFields.Add(Range("A1:A4"))
    sortfields.Order = xlDescending
}
```

```javascript
/*本示例在第一张工作表创建新的排序字段，并显示该排序字段的排序次序是否为升序。*/
function test() {
    Application.Sheets.Item(1).Activate()
    let sortfield = ActiveSheet.Sort.SortFields.Add(Range("C1:C10"), xlSortOnValues, xlAscending)
    console.log(sortfield.Order == xlAscending)
}
```
