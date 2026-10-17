---
title: "Sheets (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Workbook/%E5%B1%9E%E6%80%A7/Sheets%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: Dc4oSPJU
---

# Sheets (属性)

返回一个 Sheets 集合，它代表指定工作簿中所有工作表。 **[Sheets](./../../Sheets/Sheets 对象.html)** 对象，只读。

## 说明

在不使用对象识别符的情况下使用此属性等效于使用 `ActiveWorkbook.Sheets`。

## 示例

```javascript
/*此示例新建一张工作表，然后在第一列中列出活动工作簿中的工作表名称。*/
function test() {
    let newSheet = Application.Sheets.Add(null, null, null, xlWorksheet)
    for (let i = 1; i <= Application.Sheets.Count; i++) {
        newSheet.Cells.Item(i, 1).Value2 = Application.Sheets.Item(i).Name
    }
}
```

```javascript
/*此示例演示工作表的总数量。*/
function test() {
    console.log("工作表的总数量：" + Application.Sheets.Count)
}
```
