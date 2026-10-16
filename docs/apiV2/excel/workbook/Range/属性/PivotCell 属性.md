---
title: "PivotCell 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Range/%E5%B1%9E%E6%80%A7/PivotCell%20%E5%B1%9E%E6%80%A7.html
---
# PivotCell (属性)

返回一个 **[PivotCell](./../../PivotCell/PivotCell 对象.html)** 对象，该对象表示数据透视表中的一个单元格。

## 示例

```javascript
/*本示例确定 PivotCell 对象所在的数据透视表的名称，并通知用户。本示例假定数据透视表位于活动工作表上，并且单元格 A3 位于数据透视表上。*/
function test() {
    //Determine the name of the PivotTable the PivotCell is located in.
    console.log ("Cell A3 is located in PivotTable: " + Application.Range("A3").PivotCell.Parent)
}
```

```javascript
/*本示例显示 B3 单元格数据透视表字段的标签文本的文本和所关联的数据透视表名称*/
function test() {
    console.log("B3单元格数据透视表字段的标签文本：" + Range("B3").PivotCell.PivotField.Caption)
    console.log("B3单元格所关联的数据透视表为：" + Range("B3").PivotCell.PivotTable.Name)
}
```
