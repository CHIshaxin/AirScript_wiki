---
title: "DefaultPivotTableStyle 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Workbook/%E5%B1%9E%E6%80%A7/DefaultPivotTableStyle%20%E5%B1%9E%E6%80%A7.html
---
# DefaultPivotTableStyle (属性)

从 **[TableStyles](./../../TableStyles/TableStyles 对象.html)** 集合中指定表样式，该表样式将用作数据透视表的默认样式。可读/写。

## 示例

```javascript
/*本示例演示替换活动工作簿数据透视表的默认样式*/
function test() {
    console.log("原样式:" + ActiveWorkbook.DefaultPivotTableStyle)
    ActiveWorkbook.DefaultPivotTableStyle = "PivotStylePreset3_Accent2"
    console.log("替换为:" + ActiveWorkbook.DefaultPivotTableStyle)
}
```

```javascript
/*本示例演示向活动工作簿的 TableStyles 集合中添加 pivotTable_style1 样式，并指定 pivotTable_style1 为数据透视表默认样式*/
function test() {
    ActiveWorkbook.TableStyles.Add("pivotTable_style1")
    ActiveWorkbook.DefaultPivotTableStyle = "pivotTable_style1"
}
```
