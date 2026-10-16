---
title: "Priority (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/SortField/%E5%B1%9E%E6%80%A7/Priority%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: 2G8PCHxE
---

# Priority (属性)

指定排序字段的优先级。可读/写。

## 示例

```javascript
/*本示例在第一张工作表创建新的排序字段，并显示该排序字段是否为最主要的排序字段。*/
function test() {
    let sortfield = Application.Sheets.Item(1).Sort.SortFields.Add(Range("B1:B4"))
    console.log(sortfield.Priority == 1)
}
```

```javascript
/*本示例判断如果活动工作表 SortField 对象的数量大于2，则将活动工作表第三个排序字段设置为次要的排序字段。*/
function test() {
    let sortfields = Application.ActiveSheet.Sort.SortFields
    if (sortfields.Count > 2) {
        sortfields.Item(3).Priority = 2
    }
}
```
