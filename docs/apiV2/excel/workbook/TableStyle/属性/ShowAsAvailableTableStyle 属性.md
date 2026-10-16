---
title: "ShowAsAvailableTableStyle 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/TableStyle/%E5%B1%9E%E6%80%A7/ShowAsAvailableTableStyle%20%E5%B1%9E%E6%80%A7.html
---
# ShowAsAvailableTableStyle (属性)

返回或设置在表样式库中显示为可用的表样式。可读/写 **Boolean** 类型。

## 说明

如果为 **True**，则此样式显示在表样式库中。

即使样式已应用于表，也可以将此属性设置为 **False**。这种情况下，库中将不会显示该样式，活动单元格在该表中时，样式状态不会显示为选中。

## 示例

```javascript
/*本示例将活动工作簿上的表格样式中的最后一个样式设置为在表样式库中显示为可用的表样式。*/
function test() {
    let tableStyles = ActiveWorkbook.TableStyles
    tableStyles.Item(tableStyles.Count).ShowAsAvailableTableStyle = true
}
```

```javascript
/*本示例显示工作表 Sheet1 上第一张列表的表样式在表样式库中是否显示为可用的表样式。*/
function test() {
    let listObj = Application.Worksheets.Item("Sheet1").ListObjects.Item(1)
    console.log(listObj.TableStyle.ShowAsAvailableTableStyle)
}
```
