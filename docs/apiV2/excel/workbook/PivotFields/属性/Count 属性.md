---
title: "Count 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/PivotFields/%E5%B1%9E%E6%80%A7/Count%20%E5%B1%9E%E6%80%A7.html
---
# Count (属性)

返回一个 **Long** 值，它代表集合中对象的数量。

## 示例

```javascript
/*本示例显示第一张工作表中数据透视表的字段数量。*/
function test() {
    console.log(Worksheets.Item(1).Range("I1").PivotTable.PivotFields().Count)
}
```

```javascript
/*本示例显示活动工作表中第一张数据透视表的字段数量。*/
function test() {
    console.log(`字段数量：${ActiveSheet.PivotTables(1).PivotFields().Count}`)
}
```
