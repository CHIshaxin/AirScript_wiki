---
title: "PivotTable (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/PivotCell/%E5%B1%9E%E6%80%A7/PivotTable%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: DvFGnYFu
---

# PivotTable (属性)

返回一个 **[PivotTable](./../../PivotTable/PivotTable 对象.html)** 对象，它代表与 PivotCell 相关联的数据透视表。

## 示例

```javascript
/*此示例将包含 Sheet1 上单元格 L7 的数据透视项所在的数据透视表的当前页设置为名为“5”的页。*/
function test() {
    let pvtTable = Worksheets.Item("Sheet1").Range("L7").PivotCell.PivotTable
    pvtTable.PivotFields("id").CurrentPage = "5"
}
```

```javascript
/*此示例将包含活动单元格的数据透视项所在的数据透视表中名为“张三”的页设置为数据透视表的当前页。*/
function test() {
    ActiveCell.PivotCell.PivotTable.PivotFields().Item("name").CurrentPage = "张三"
}
```
