---
title: "Visible 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/PivotItem/%E5%B1%9E%E6%80%A7/Visible%20%E5%B1%9E%E6%80%A7.html
---
# Visible (属性)

返回或设置一个 **Boolean** 值，它确定对象是否可见。可读写。

## 说明

如果数据透视表中的某一项在当前是可见的，则该项的 **Visible** 属性为 **True**。

## 示例

```javascript
/*本示例显示活动工作表中数据透视表的字段“id”的第二个数据项是否可见。*/
function test() {
    let pvtItem = ActiveSheet.Range("I1").PivotTable.PivotFields("id").PivotItems(2)
    console.log(pvtItem.Visible)
}
```

```javascript
/*本示例将工作表 Sheet1 中第一张数据透视表的字段“name”中数据项“张三”设置为不可见。*/
function test() {
    let pvtItem = Worksheets.Item("Sheet1").PivotTables(1).PivotFields("name").PivotItems("张三")
    pvtItem.Visible = false
}
```
