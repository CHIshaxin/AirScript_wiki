---
title: "PivotTable (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Range/%E5%B1%9E%E6%80%A7/PivotTable%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: abZx4tlS
---

# PivotTable (属性)

返回一个 **[PivotTable](./../../PivotTable/PivotTable 对象.html)** 对象，它代表指定区域左上角所在的数据透视表。

## 示例

```javascript
/*此示例将 Sheet1 上的数据透视表的当前页设置为名为“Canada”的页。*/
function test() {
    let pvtTable = Worksheets.Item("Sheet1").Range("C3").PivotTable
    pvtTable.PivotFields("Country").CurrentPage = "Canada"
}
```

```javascript
/*本示例显示 C3 单元格所属的数据透视表，并将其清除。*/
function test() {
    console.log("C3单元格所属的数据透视表为：" + Range("C3").PivotTable.Name)
    Range("C3").PivotTable.ClearTable()
}
```
