---
title: "Type (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Worksheet/%E5%B1%9E%E6%80%A7/Type%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: DHlSFlTl
---

# Type (属性)

返回一个代表工作表类型的 XlSheetType值。

## 示例

```javascript
/*本示例显示活动工作表类型的 XlSheetType 值。*/
function test() {
    console.log(`活动工作表的类型值为 ${Application.ActiveSheet.Type}`)
}
```

```javascript
/*本示例判断活动工作簿中每一张工作表类型，如果是工作表或图表工作表则通知用户。*/
function test() {
    for (let i = 1; i <= Sheets.Count; i++) {
        let sheet = Application.Worksheets.Item(i)
        if (sheet.Type == xlWorksheet) {
            console.log(`工作表 ${sheet.Name} 类型为工作表`)
        } else if (sheet.Type == xlChart) {
            console.log(`工作表 ${sheet.Name} 类型为图表工作表`)
        }
    }
}
```
