---
title: "Delete (方法)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/PivotField/%E6%96%B9%E6%B3%95/Delete%20%E6%96%B9%E6%B3%95.html
updated: 1789700916000
hash: rZGXjyFj
---

# Delete (方法)

删除对象。

## 示例

```javascript
/*本示例删除活动工作表中第一个数据透视表的字段“学号”。*/
function test() {
    ActiveSheet.PivotTables(1).PivotFields("学号").Delete()
}
```

```javascript
/*本示例删除工作表 Sheet1 中第一个数据透视表的字段“期中成绩”。*/
function test() {
    let pvtField = Worksheets.Item("Sheet1").PivotTables(1).PivotFields("期中成绩")
    pvtField.Delete()
}
```
