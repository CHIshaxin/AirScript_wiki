---
title: "CustomOrder (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/SortField/%E5%B1%9E%E6%80%A7/CustomOrder%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: maDw6wDZ
---

# CustomOrder (属性)

指定对字段进行排序的自定义次序。可读/写 **Variant** 类型。

## 示例

```javascript
/*本示例在活动工作表设置自定义排序。*/
function test() {
    Application.ActiveSheet.Sort.SortFields.Item(1).CustomOrder = "A,B,C,D,E,F,G,H"
}
```

```javascript
/*本示例在第一张工作表创建新的排序字段，并显示第一张工作表的自定义排序。*/
function test() {
    Application.Sheets.Item(1).Sort.SortFields.Add(Range("A1:A20"), xlSortOnValues, xlAscending, "alpha,bravo,charlie,delta")
    console.log(Application.Sheets.Item(1).Sort.SortFields.Item(1).CustomOrder)
}
```
