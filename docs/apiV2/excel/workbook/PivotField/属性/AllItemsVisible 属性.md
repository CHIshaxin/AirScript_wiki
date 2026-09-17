---
title: "AllItemsVisible 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/PivotField/%E5%B1%9E%E6%80%A7/AllItemsVisible%20%E5%B1%9E%E6%80%A7.html
---
# AllItemsVisible (属性)

用于检索指示是否对透视字段应用任何手动筛选的 Boolean 值。只读。

## 说明

此属性提供了一种简单方式，便于检查是否对透视字段或多维数据集字段应用手动筛选。

对于 OLAP 数据透视表，此属性仅适用于 **CubeField** 对象。如果尝试从 OLAP 数据透视表中的 **[PivotField](./../../PivotField/PivotField 对象.html)** 对象获取此属性或设置此属性，将返回运行时错误。

对于数据透视表，此属性适用于 **[PivotField](./../../PivotField/PivotField 对象.html)** 对象。

默认值为 **True**。未应用手动筛选（无论 **IncludeNewItemsInFilter** 属性为 **True** 还是为 **False**）时，此属性将自动设置为 **True**。应用任何手动筛选（无论 **IncludeNewItemsInFilter** 属性为 **True** 还是为 **False**）时，此属性将自动设置为 **False**。

此属性直接反映透视字段或多维数据集字段的筛选下拉列表中 **“全选”** 复选框的状态。

## 示例

```javascript
/*本示例显示工作表 Sheet1 中第一张数据透视表的“id”字段是否未应用手动筛选。*/
function test() {
    console.log(Worksheets.Item("Sheet1").PivotTables(1).PivotFields("id").AllItemsVisible)
}
```

```javascript
/*本示例判断活动工作表中第一张数据透视表的“grades”字段是否应用手动筛选，并通知用户。*/
function test() {
    let pvtField = ActiveSheet.PivotTables(1).PivotFields("grades")
    if (pvtField.AllItemsVisible) {
        console.log("字段未应用手动筛选")
    } else {
        console.log("字段应用手动筛选")
    }
}
```
