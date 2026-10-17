---
title: "Protection (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Worksheet/%E5%B1%9E%E6%80%A7/Protection%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: u-vkYGAl
---

# Protection (属性)

返回一个 Protection 对象，该对象表示工作表的保护选项。

## 示例

```javascript
/*本示例对活动工作表进行保护，并判断是否能在受保护的工作表中插入列，然后将此状态通知用户。*/
function test() {
    Application.ActiveSheet.Protect()
    // Check the ability to insert columns on a protected sheet.
    // Notify the user of this status.
    if (Application.ActiveSheet.Protection.AllowInsertingColumns == true) {
        console.log("The insertion of columns is allowed on this protected worksheet.")
    } else {
        console.log("The insertion of columns is not allowed on this protected worksheet.")
    }
}
```

```javascript
/*本示例允许用户对受保护的工作表 Sheet1 上的单元格进行格式设置，并通知用户。*/
function test() {
    Worksheets.Item("Sheet1").Protect(null, null, null, null, null, true)
    if (Worksheets.Item("Sheet1").Protection.AllowFormattingCells) {
        console.log("可以在此受保护的工作表 Sheet1 上设置单元格的格式。")
    } else {
        console.log("不能在此受保护的工作表 Sheet1 上设置单元格的格式。")
    }
}
```
