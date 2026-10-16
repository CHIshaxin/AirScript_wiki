---
title: "Version (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/PivotTable/%E5%B1%9E%E6%80%A7/Version%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: BGyaIYRk
---

# Version (属性)

返回一个 **[XlPivotTableVersionList](./../../枚举/XlPivotTableVersionList 枚举.html)** 值，它代表 ET 版本号。

## 示例

```javascript
/*本示例显示活动工作表中第一张数据透视表的版本是否为“xlPivotTableVersion10”。*/
function test() {
    console.log(ActiveSheet.PivotTables(1).Version == xlPivotTableVersion10)
}
```

```javascript
/*本示例判断工作表 Sheet1 中数据透视表的版本是否为“xlPivotTableVersion11”，并通知用户。*/
function test() {
    let pvtTable = Worksheets.Item("Sheet1").Range("I1").PivotTable
    if (pvtTable.Version == xlPivotTableVersion11) {
        console.log("版本号是“xlPivotTableVersion11”")
    } else {
        console.log("版本号不是“xlPivotTableVersion11”")
    }
}
```
