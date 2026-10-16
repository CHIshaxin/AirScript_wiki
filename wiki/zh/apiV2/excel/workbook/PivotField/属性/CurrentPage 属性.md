---
title: "CurrentPage (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/PivotField/%E5%B1%9E%E6%80%A7/CurrentPage%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: C64hdAwu
---

# CurrentPage (属性)

返回或设置页字段的当前页显示（仅对页字段有效）。**PivotItem** 类型，可读写。

## 示例

```javascript
/*本示例将 Sheet1 上数据透视表中字段“id”的当前页名称返回到字符串变量 `strPgName` 中并显示。*/
function test() {
    let pvtTable = Worksheets.Item("Sheet1").Range("I2").PivotTable
    let strPgName = pvtTable.PivotFields("id").CurrentPage
    console.log(strPgName)
}
```

```javascript
/*本示例将活动工作表第一张数据透视表的页字段“id”显示的当前页设置为“5”。*/
function test() {
    ActiveSheet.PivotTables(1).PivotFields("id").CurrentPage = "5"
}
```
