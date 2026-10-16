---
title: "DataType (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/PivotField/%E5%B1%9E%E6%80%A7/DataType%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: 6mBBbAWn
---

# DataType (属性)

返回一个 **[XlPivotFieldDataType](./../../枚举/XlPivotFieldDataType 枚举.html)** 值，它代表数据透视表字段中的数据类型。

## 示例

```javascript
/*本示例显示工作表 Sheet1 中数据透视表的字段“name”的数据类型。*/
function test() {
    let pvtTable = Worksheets.Item("Sheet1").Range("I1").PivotTable
    switch (pvtTable.PivotFields("name").DataType) {
        case xlText:
            console.log("The field contains text data")
            break
        case xlNumber:
            console.log("The field contains numeric data")
            break
        case xlDate:
            console.log("The field contains date data")
    }
}
```

```javascript
/*本示例显示活动工作表中第一个数据透视表的字段“grades”的数据类型是否为“xlNumber”。*/
function test() {
    console.log(ActiveSheet.PivotTables(1).PivotFields("grades").DataType == xlNumber)
}
```
