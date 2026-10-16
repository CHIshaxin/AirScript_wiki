---
title: "GrandTotalName (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/PivotTable/%E5%B1%9E%E6%80%A7/GrandTotalName%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: BZY2fiub
---

# GrandTotalName (属性)

返回或设置显示在指定数据透视表的总计列或行标题中的文本串标志。默认值为字符串“Grand Total”。**String** 类型，可读写。

## 示例

```javascript
/*本示例将活动工作表的第二个数据透视表中总计标题的标志设置为“Regional Total”。*/
function test() {
    Application.ActiveSheet.PivotTables("PivotTable2").GrandTotalName = "Regional Total"
}
```

```javascript
/*本示例将第一张工作表中第二张数据透视表的汇总列标题设置为总和，并弹窗显示该属性。*/
function test() {
    Worksheets.Item(1).PivotTables(2).GrandTotalName = "总和"
    console.log(Worksheets.Item(1).PivotTables(2).GrandTotalName)
}
```
