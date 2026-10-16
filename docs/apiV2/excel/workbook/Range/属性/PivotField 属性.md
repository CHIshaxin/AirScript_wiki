---
title: "PivotField 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Range/%E5%B1%9E%E6%80%A7/PivotField%20%E5%B1%9E%E6%80%A7.html
---
# PivotField (属性)

返回一个 **[PivotField](./../../PivotField/PivotField 对象.html)** 对象，它代表指定区域左上角所在的数据透视表字段。

## 示例

```javascript
/*此示例显示包含活动单元格的数据透视表字段的名称。*/
function test() {
    Worksheets.Item("Sheet1").Activate()
    console.log("The active cell is in the field " + ActiveCell.PivotField.Name)
}
```

```javascript
/*本示例将 A3 单元格所在的数据透视表字段名设置为“字段一”，并显示*/
function test() {
    Range("A3").PivotField.Value = "字段一"
    console.log(Range("A3").PivotField.Value)
}
```
