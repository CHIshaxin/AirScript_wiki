---
title: "AllowSorting (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Protection/%E5%B1%9E%E6%80%A7/AllowSorting%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: BJYAjPEM
---

# AllowSorting (属性)

如果允许在受保护的工作表上使用排序选项，则返回 **True**。**Boolean** 类型，只读。

## 说明

在受保护的工作表中，只能对未锁定或未保护的单元格进行排序。

可以使用 **Protect** 方法参数设置 **AllowSorting** 属性。

## 示例

```javascript
/*本示例允许用户对受保护的工作表上未锁定或未保护的单元格进行排序，并通知用户。*/
function test() {
    Application.ActiveSheet.Unprotect()
    // Unlock cells A1 through B5.
    Range("A1:B5").Locked = false
    // Allow sorting to be performed on the protected worksheet.
    if (Application.ActiveSheet.Protection.AllowSorting == false) {
        Application.ActiveSheet.Protect(null, null, null, null, null, null, null, null, null, null, null, null, null, true)
    }
    console.log("For cells A1 through B5, sorting can be performed on the protected worksheet.")
}
```

```javascript
/*本示例显示是否允许在受保护的第二张工作表上使用排序选项。*/
function test() {
    console.log(Worksheets.Item(2).Protection.AllowSorting)
}
```
