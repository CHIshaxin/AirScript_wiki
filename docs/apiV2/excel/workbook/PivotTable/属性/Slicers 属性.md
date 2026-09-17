---
title: "Slicers 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/PivotTable/%E5%B1%9E%E6%80%A7/Slicers%20%E5%B1%9E%E6%80%A7.html
---
# Slicers (属性)

返回指定数据透视表的 **[Slicers](./../../Slicers/Slicers 对象.html)** 集合。只读

## 说明

切片器可用于筛选数据透视表中的数据。 **[Slicers](./../../Slicers/Slicers 对象.html)** 集合包含的 **[Slicer](./../../Slicer/Slicer 对象.html)** 对象表示与该指定数据透视表关联的切片器。

## 示例

```javascript
/*本示例删除数据透视表中所有的切片器*/
function test() {
    for (i = 1; i <= Worksheets.Item(1).PivotTables(1).Slicers.Count; i++) {
        Worksheets.Item(1).PivotTables(1).Slicers.Item(i).Delete()
    }
}
```

```javascript
/*本示例将数据透视表中第一个切片器的标题设置为“总量”*/
function test() {
    ActiveSheet.PivotTables(1).Slicers.Item(1).Caption = "总量"
}
```
