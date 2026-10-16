---
title: "LocationInTable (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Range/%E5%B1%9E%E6%80%A7/LocationInTable%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: DnI2LdQF
---

# LocationInTable (属性)

返回一个常量，该常量描述包含指定区域左上角部分的 **PivotTable** 部分。可为以下 **[XlLocationInTable](./../../枚举/XlLocationInTable 枚举.html)** 常量之一。**Long** 类型，只读。

## 说明

| **[XlLocationInTable](./../../枚举/XlLocationInTable 枚举.html)** 可为以下 **[XlLocationInTable](./../../枚举/XlLocationInTable 枚举.html)** 常量之一。 |
| --- |
| **xlRowHeader** |
| **xlColumnHeader** |
| **xlPageHeader** |
| **xlDataHeader** |
| **xlRowItem** |
| **xlColumnItem** |
| **xlPageItem** |
| **xlDataItem** |
| **xlTableBody** |

## 示例

```javascript
/*本示例显示一个消息框，用以描述数据透视表中活动单元格所在的位置。*/
function test() {
    Worksheets.Item("Sheet1").Activate()
    switch (Selection.LocationInTable) {
        case xlRowHeader :
            console.log("Active cell is part of a row header")
            break
        case xlColumnHeader :
            console.log("Active cell is part of a column header")
            break
        case xlPageHeader :
            console.log("Active cell is part of a page header")
            break
        case xlDataHeader :
            console.log("Active cell is part of a data header")
            break
        case xlRowItem :
            console.log("Active cell is part of a row item")
            break
        case xlColumnItem :
            console.log("Active cell is part of a column item")
            break
        case xlPageItem :
            console.log("Active cell is part of a page item")
            break
        case xlDataItem :
            console.log("Active cell is part of a data item")
            break
        case xlTableBody :
            console.log("Active cell is part of the table body")
            break
    }
}
```

```javascript
/*本示例判断 A2 单元格是否为数据透视表的正文*/
function test() {
    if (Range("A2").LocationInTable == xlTableBody) {
        console.log("A2单元格为数据透视表的正文")
    }
}
```
