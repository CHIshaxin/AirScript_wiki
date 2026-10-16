---
title: "Name 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/PivotTable/%E5%B1%9E%E6%80%A7/Name%20%E5%B1%9E%E6%80%A7.html
---
# Name (属性)

返回或设置一个 **String** 值，它代表对象的名称。

## 示例

```javascript
/*本示例显示活动工作表中第一张数据透视表的名称*/
function test() {
    console.log(ActiveSheet.PivotTables(1).Name)
}
```

```javascript
/*本示例将工作表 Sheet1 中第一张数据透视表的名称设置为数据透视表1，并弹窗显示此属性*/
function test() {
    Worksheets.Item("Sheet1").PivotTables(1).Name = "数据透视表1"
    console.log(Worksheets.Item("Sheet1").PivotTables(1).Name)
}
```
