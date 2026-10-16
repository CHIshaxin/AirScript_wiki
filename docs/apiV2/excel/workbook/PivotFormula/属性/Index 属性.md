---
title: "Index 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/PivotFormula/%E5%B1%9E%E6%80%A7/Index%20%E5%B1%9E%E6%80%A7.html
---
# Index (属性)

返回或设置一个 **Long** 值，它代表 **[PivotFormula](./../../PivotFormula/PivotFormula 对象.html)** 对象在 **[PivotFormulas](./../../PivotFormulas/PivotFormulas 对象.html)** 集合中的索引号。

## 示例

```javascript
/*本示例显示第一张工作表中数据透视表名为“公式1”的公式的索引。*/
function test() {
    let pvtFormula = Worksheets.Item(1).Range("I1").PivotTable.PivotFormulas("公式1")
    console.log(pvtFormula.Index)
}
```

```javascript
/*本示例将第一张工作表的第一张数据透视表中第二个公式的索引设置为 1 并显示。*/
function test() {
    let pvtFormula = Worksheets.Item(1).PivotTables(1).PivotFormulas.Item(2)
    pvtFormula.Index = 1
    console.log(pvtFormula.Index)
}
```
