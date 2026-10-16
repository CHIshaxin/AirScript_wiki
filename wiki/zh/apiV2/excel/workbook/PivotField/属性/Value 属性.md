---
title: "Value (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/PivotField/%E5%B1%9E%E6%80%A7/Value%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: CX_O5xpz
---

# Value (属性)

返回或设置一个 **String** 值，它代表数据透视表中指定的字段的名称。

## 示例

```javascript
/*本示例显示活动工作表中第一张数据透视表的第一个字段的名称。*/
function test() {
    console.log(ActiveSheet.PivotTables(1).PivotFields(1).Value)
}
```

```javascript
/*本示例将第一张工作表中数据透视表的字段“id”的名称修改为“学号”。*/
function test() {
    let pvtField = Worksheets.Item(1).Range("I1").PivotTable.PivotFields("id")
    pvtField.Value = "学号"
}
```
