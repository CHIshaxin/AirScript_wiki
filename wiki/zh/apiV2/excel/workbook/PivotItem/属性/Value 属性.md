---
title: "Value (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/PivotItem/%E5%B1%9E%E6%80%A7/Value%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: CTmJ2BF7
---

# Value (属性)

返回或设置一个 **String** 值，它代表数据透视表中指定的数据项的名称。

## 示例

```javascript
/*本示例显示活动工作表中数据透视表的字段“name”的第二个数据项的名称。*/
function test() {
    let pvtItem = ActiveSheet.Range("I1").PivotTable.PivotFields("name").PivotItems(2)
    console.log(pvtItem.Value)
}
```

```javascript
/*本示例将工作表 Sheet1 中第一张数据透视表的字段“name”中数据项“张三”的名称设置为“张三三”。*/
function test() {
    let pvtItem = Worksheets.Item("Sheet1").PivotTables(1).PivotFields("name").PivotItems("张三")
    pvtItem.Value = "张三三"
}
```
