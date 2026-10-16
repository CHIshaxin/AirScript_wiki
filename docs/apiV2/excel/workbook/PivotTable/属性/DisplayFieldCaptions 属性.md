---
title: "DisplayFieldCaptions 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/PivotTable/%E5%B1%9E%E6%80%A7/DisplayFieldCaptions%20%E5%B1%9E%E6%80%A7.html
---
# DisplayFieldCaptions (属性)

控制是否在网格中显示筛选器按钮以及行和列的透视字段标题。可读/写。

## 说明

默认值为 **True**。

## 示例

```javascript
/*本示例隐藏第一张工作表中第一张数据透视表筛选器按钮以及行和列的透视字段标题*/
function test() {
    Worksheets.Item(1).PivotTables(1).DisplayFieldCaptions = false
}
```

```javascript
/*本示例显示活动工作表中数据透视表1的 DisplayFieldCaptions 的数据类型，并设置为 true*/
function test() {
    console.log(typeof ActiveSheet.PivotTables("数据透视表1").DisplayFieldCaptions)
    ActiveSheet.PivotTables("数据透视表1").DisplayFieldCaptions = true
}
```
