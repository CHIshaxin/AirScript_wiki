---
title: "EnableOutlining (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Worksheet/%E5%B1%9E%E6%80%A7/EnableOutlining%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: DuB9fI37
---

# EnableOutlining (属性)

如果当仅限于用户界面保护处于打开状态时，启用分级显示符号，则该属性值为 **True**。**Boolean** 类型，可读写。

## 示例

```javascript
/*本示例在受保护的工作表上启用分级显示符号。*/
function test() {
    Application.ActiveSheet.EnableOutlining = true
    Application.ActiveSheet.Protect(null, null, true, null, true)
}
```

```javascript
/*本示例判断如果第一张工作表启用分级显示符号，则将分级显示符号关闭。*/
function test() {
    if (Application.Sheets.Item(1).EnableOutlining) {
        Application.Sheets.Item(1).EnableOutlining = false
    }
}
```
