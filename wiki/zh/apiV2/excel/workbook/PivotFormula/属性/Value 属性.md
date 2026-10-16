---
title: "Value (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/PivotFormula/%E5%B1%9E%E6%80%A7/Value%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: F1lPv-p5
---

# Value (属性)

返回或设置一个 **String** 值，它代表数据透视表中指定的公式的名称。

## 示例

```javascript
/*本示例显示活动工作表中第一张数据透视表第一个公式的名称。*/
function test() {
    let pvtFormula = ActiveSheet.PivotTables(1).PivotFormulas(1)
    console.log(pvtFormula.Value)
}
```

```javascript
/*本示例显示第一张工作表中数据透视表第二个公式的名称。*/
function test() {
    let pvtFormula = Worksheets.Item(1).Range("I1").PivotTable.PivotFormulas(2)
    console.log(pvtFormula.Value)
}
```
