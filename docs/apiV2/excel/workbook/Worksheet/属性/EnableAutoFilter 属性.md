---
title: "EnableAutoFilter 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Worksheet/%E5%B1%9E%E6%80%A7/EnableAutoFilter%20%E5%B1%9E%E6%80%A7.html
---
# EnableAutoFilter (属性)

如果当仅限于用户界面保护处于打开状态时，启用自动筛选箭头，则该属性值为 **True**。**Boolean** 类型，可读写。

## 示例

```javascript
/*本示例启用受保护的工作表上的自动筛选箭头。*/
function test() {
    Application.ActiveSheet.EnableAutoFilter = true
    Application.ActiveSheet.Protect(null, null, true, null, true)
}
```

```javascript
/*本示例判断活动工作簿上每一张工作表是否启用自动筛选箭头，并通知用户。*/
function test() {
    for (let i = 1; i <= Sheets.Count; i++) {
        let sheet = Application.Sheets.Item(i)
        if (sheet.EnableAutoFilter) {
            console.log(`工作表 ${sheet.Name} 启用自动筛选箭头`)
        } else {
            console.log(`工作表 ${sheet.Name} 未启用自动筛选箭头`)
        }
    }
}
```
