---
title: "Sheets (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Application/%E5%B1%9E%E6%80%A7/Sheets%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: DvaXBIgo
---

# Sheets (属性)

返回一个 **[Sheets](./../../Sheets/Sheets 对象.html)** 集合，它代表活动工作簿中所有的工作表。 **[Sheets](./../../Sheets/Sheets 对象.html)** 对象，只读。

## 说明

在不使用对象识别符的情况下使用此属性等效于使用 `ActiveWorkbook.Sheets`。

## 示例

```javascript
/*此示例新建一张工作表，然后在第一列中列出活动工作簿中的工作表名称。*/
function test() {
    let newSheet = Application.Sheets.Add(undefined, undefined, undefined, xlWorksheet)
    for (let i = 1; i <= Application.Sheets.Count; i++) {
        newSheet.Cells(i, 1).Value(undefined, Application.Sheets(i).Name)
    }
}
```

```javascript
/*此示例演示复制所有工作表在 Sheet1 之前*/
function test() {
    Application.Sheets.Copy(Application.Sheets.Item('Sheet1'))
}
```
