---
title: "TableStyle2 (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/PivotTable/%E5%B1%9E%E6%80%A7/TableStyle2%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: DniqwFSq
---

# TableStyle2 (属性)

**TableStyle2** 属性指定当前应用于数据透视表的数据透视表样式。可读/写。

## 说明

由于存在名为 **TableStyle** 的属性，因此该属性称为 TableStyle2。

## 示例

```javascript
/*本示例设置活动工作表第一张数据透视表的样式。*/
function test() {
    ActiveSheet.PivotTables(1).TableStyle2 = "PivotStyleLight17"
}
```

```javascript
/*本示例设置工作表 Sheet1 中数据透视表的样式为“PivotStyleLight6”。*/
function test() {
    Worksheets.Item("Sheet1").Range("I1").PivotTable.TableStyle2 = "PivotStyleLight6"
}
```
