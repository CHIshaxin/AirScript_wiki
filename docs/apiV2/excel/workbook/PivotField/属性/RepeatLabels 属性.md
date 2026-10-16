---
title: "RepeatLabels 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/PivotField/%E5%B1%9E%E6%80%A7/RepeatLabels%20%E5%B1%9E%E6%80%A7.html
---
# RepeatLabels (属性)

返回或设置在数据透视表中是否对指定的透视字段重复项目标签。可读/写

## 说明

如果对指定的透视字段重复项目标签，则为 **True**；否则为 **False**。

**RepeatLabels**属性的设置对应于数据透视表中的字段的 **“字段设置”** 对话框 **“布局和打印”** 选项卡上的 **“重复项目标签”** 复选框。

若要指定是否在单次操作中对数据透视表内的所有透视字段重复项目标签，请使用 **RepeatAllLabels** 方法。

## 示例

```javascript
/*本示例显示活动工作表中第一张数据透视表的字段“name”是否重复项目标签。*/
function test() {
    console.log(ActiveSheet.PivotTables(1).PivotFields("name").RepeatLabels)
}
```

```javascript
/*本示例将第一张工作表中数据透视表的字段“score”设置为重复项目标签。*/
function test() {
    let pvtField = Worksheets.Item(1).Range("I1").PivotTable.PivotFields("score")
    pvtField.RepeatLabels = true
}
```
