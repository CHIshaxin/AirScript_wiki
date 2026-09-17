---
title: "SortFields 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Sort/%E5%B1%9E%E6%80%A7/SortFields%20%E5%B1%9E%E6%80%A7.html
---
# SortFields (属性)

可用来在工作簿、列表和自动筛选上存储排序状态。只读 **SortFields**类型。

## 示例

```javascript
/*本示例在活动工作表创建新的排序字段。*/
function test() {
    Application.ActiveSheet.Sort.SortFields.Add(Range("B1:B10"))
}
```

```javascript
/*本示例在第一张工作表添加两个排序字段，并显示排序字段集合的数量。*/
function test() {
    let sortfields = Application.Sheets.Item(1).Sort.SortFields
    sortfields.Add(Range("A1:A5"))
    sortfields.Add(Range("B1:B5"))
    console.log(sortfields.Count)
}
```
