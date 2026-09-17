---
title: "ShowAsAvailablePivotTableStyle 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/TableStyle/%E5%B1%9E%E6%80%A7/ShowAsAvailablePivotTableStyle%20%E5%B1%9E%E6%80%A7.html
---
# ShowAsAvailablePivotTableStyle (属性)

设置或返回一个值，该值表示是否在数据透视表样式库中显示样式，可读/写 **Boolean** 类型。

## 说明

如果在数据透视表样式库中显示样式，则返回 **True**。

即使已将样式应用于表格或数据透视表，用户仍可以将 **ShowAsAvailableTableStyle** 或 **ShowAsAvailablePivotTableStyle** 属性设置为 **False**。在这种情况下，当活动单元格位于表格或数据透视表中时，库将不会显示样式并且没有样式显示为选中的样式。

## 示例

```javascript
/*本示例显示工作表 Sheet1 上第一张列表的表样式是否在数据透视表样式库中显示样式。*/
function test() {
    let listObj = Application.Worksheets.Item("Sheet1").ListObjects.Item(1)
    console.log(listObj.TableStyle.ShowAsAvailablePivotTableStyle)
}
```

```javascript
/*本示例将活动工作簿上的表格样式中的最后一个样式设置为在数据透视表样式库中显示样式。*/
function test() {
    let tableStyles = ActiveWorkbook.TableStyles
    tableStyles.Item(tableStyles.Count).ShowAsAvailablePivotTableStyle = true
}
```
