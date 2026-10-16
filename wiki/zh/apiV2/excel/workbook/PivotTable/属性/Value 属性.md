---
title: "Value (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/PivotTable/%E5%B1%9E%E6%80%A7/Value%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: C-ttYTWa
---

# Value (属性)

返回或设置一个 **String** 值，它代表数据透视表的名称。

## 示例

```javascript
/*本示例显示第一张数据表的名称*/
function test() {
    console.log(Worksheets.Item(1).PivotTables(1).Value)
}
```

```javascript
/*本示例将活动工作表中第一张数据透视表的名称设置为“总结”，并弹窗显示名称*/
function test() {
    ActiveSheet.PivotTables(1).Value = "总结"
    console.log(ActiveSheet.PivotTables(1).Value)
}
```
