---
title: "Count (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/SortFields/%E5%B1%9E%E6%80%A7/Count%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: BaNXWv3U
---

# Count (属性)

返回集合中对象的数目。只读 **Long** 类型。

## 示例

```javascript
/*本示例显示活动工作表排序字段集合的数量。*/
function test() {
    console.log(Application.ActiveSheet.Sort.SortFields.Count)
}
```

```javascript
/*本示例在工作表Sheet1添加两个排序字段，并显示排序字段集合的数量。*/
function test() {
    let sort = Application.Sheets.Item("Sheet1").Sort.SortFields
    sort.Add(Range("A1:D1"))
    sort.Add(Range("A2:D2"))
    console.log(sort.Count)
}
```
