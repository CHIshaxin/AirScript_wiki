---
title: "InGridDropZones 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/PivotTable/%E5%B1%9E%E6%80%A7/InGridDropZones%20%E5%B1%9E%E6%80%A7.html
---
# InGridDropZones (属性)

此属性用于为 **[PivotTable](./../../PivotTable/PivotTable 对象.html)** 对象切换网格中的拖放区域。在一些情况下，它还会影响数据透视表的布局。可读/写 **Boolean** 类型。

## 说明

当 **InGridDropZones** 属性设置为 **True** 时，存在网格中的拖放区域。当此属性设置为 **False** 时，不存在网格中的拖放区域。数据透视表的布局也会随此属性一起改变。

## 示例

```javascript
/*本示例开启数据透视表网格中的拖放区域*/
function test() {
    Worksheets.Item(1).PivotTables(1).InGridDropZones = true
}
```

```javascript
/*本示例显示数据透视表 InGridDropZones 属性的数据类型，并设置为 false。*/
function test() {
    console.log(typeof ActiveSheet.PivotTables(1).InGridDropZones)
    ActiveSheet.PivotTables(1).InGridDropZones = false
}
```
