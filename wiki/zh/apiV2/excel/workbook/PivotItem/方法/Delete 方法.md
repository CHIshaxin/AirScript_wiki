---
title: "Delete (方法)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/PivotItem/%E6%96%B9%E6%B3%95/Delete%20%E6%96%B9%E6%B3%95.html
updated: 1789700916000
hash: DlpEEyh3
---

# Delete (方法)

删除对象。

## 示例

```javascript
/*本示例删除活动工作表中数据透视表的“name”字段中第九个数据项。*/
function test() {
    ActiveSheet.Range("I1").PivotTable.PivotFields("name").PivotItems(9).Delete()
}
```

```javascript
/*本示例删除工作表 Sheet1 中第一张数据透视表的第一个字段中名为“公式2”的数据项。*/
function test() {
    let pvtItem = Worksheets.Item("Sheet1").PivotTables(1).PivotFields(1).PivotItems("公式2")
    pvtItem.Delete()
}
```
