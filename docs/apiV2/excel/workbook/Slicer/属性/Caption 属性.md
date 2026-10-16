---
title: "Caption 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Slicer/%E5%B1%9E%E6%80%A7/Caption%20%E5%B1%9E%E6%80%A7.html
---
# Caption (属性)

返回或设置指定切片器的标题。可读/写。

## 说明

为 **Caption** 属性指定的值显示在切片器的标题中。如果未指定 **Caption** 属性，ET 将使用基础 **[PivotField](./../../PivotField/PivotField 对象.html)** 对象的 **Caption** 属性值。

## 返回值

String

## 示例

```javascript
/*本示例显示活动工作表上数据透视表的第三个切片器的标题。*/
function test() {
    let slicer = ActiveSheet.Range("I1").PivotTable.Slicers.Item(3)
    console.log(slicer.Caption)
}
```

```javascript
/*本示例将第一张工作簿上第三个切片器缓存中名为“id”切片器的标题设置为“nameid”。*/
function test() {
    Workbooks.Item(1).SlicerCaches(3).Slicers.Item("id").Caption = "nameid"
}
```
