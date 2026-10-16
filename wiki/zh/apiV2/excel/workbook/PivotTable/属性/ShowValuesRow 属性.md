---
title: "ShowValuesRow (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/PivotTable/%E5%B1%9E%E6%80%A7/ShowValuesRow%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: C-8ZMEau
---

# ShowValuesRow (属性)

返回或设置是否显示值行。可读/写

## 说明

此属性对应于**“数据透视表选项”**对话框的**“显示”**选项卡上的**“显示值行”**复选框的设置。

## 示例

```javascript
/*本示例显示数据透视表的值行*/
function test() {
    Worksheets.Item(1).PivotTables(1).ShowValuesRow = true
}
```

```javascript
/*本示例显示数据透视表 ShowValuesRow 属性的数据类型，并设置为 false*/
function test() {
    console.log(typeof ActiveSheet.PivotTables(1).ShowValuesRow)
    ActiveSheet.PivotTables(1).ShowValuesRow = false
}
```
