---
title: "Delete (方法)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/PivotFormula/%E6%96%B9%E6%B3%95/Delete%20%E6%96%B9%E6%B3%95.html
updated: 1789700916000
hash: yFdsiFUM
---

# Delete (方法)

删除对象。

## 示例

```javascript
/*本示例删除第一张工作表中数据透视表的第一个公式。*/
function test() {
    let pvtFormula = Worksheets.Item(1).Range("I1").PivotTable.PivotFormulas(1)
    pvtFormula.Delete()
}
```

```javascript
/*本示例删除活动工作表的第一张数据透视表中名为“示例数据”的公式。*/
function test() {
    let pvtFormula = ActiveSheet.PivotTables(1).PivotFormulas("示例数据")
    pvtFormula.Delete()
}
```
