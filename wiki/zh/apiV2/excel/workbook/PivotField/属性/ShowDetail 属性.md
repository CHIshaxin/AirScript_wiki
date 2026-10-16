---
title: "ShowDetail (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/PivotField/%E5%B1%9E%E6%80%A7/ShowDetail%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: CBMTq8_2
---

# ShowDetail (属性)

获取或设置指定的 **PivotField** 是否显示明细数据。可读/写 **Boolean** 类型。

## 示例

```javascript
/*本示例将第一张工作表中第一张数据透视表的字段“name”的详细信息设置为可见。*/
function test() {
    Worksheets.Item(1).PivotTables(1).PivotFields("name").ShowDetail = true
}
```

```javascript
/*本示例将活动工作表中数据透视表的第一个字段的详细信息设置为不可见。*/
function test() {
    ActiveSheet.Range("I1").PivotTable.PivotFields(1).ShowDetail = false
}
```
