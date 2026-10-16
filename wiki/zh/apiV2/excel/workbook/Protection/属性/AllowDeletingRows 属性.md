---
title: "AllowDeletingRows (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Protection/%E5%B1%9E%E6%80%A7/AllowDeletingRows%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: _j1XN_8n
---

# AllowDeletingRows (属性)

如果允许删除受保护工作表上的行，则返回 **True**。**Boolean** 类型，只读。

## 说明

可以使用 **Protect** 方法参数设置 **AllowDeletingRows** 属性。

对于受保护的工作表，必须取消对包含要删除的单元格的行的锁定。

## 示例

```javascript
/*本示例取消对受保护的工作表上的第 1 行的锁定，然后允许用户删除第 1 行并通知用户。*/
function test() {
    Application.ActiveSheet.Unprotect()
    // Unlock row 1.
    Rows.Item("1:1").Locked = false
    // Allow row 1 to be deleted on a protected worksheet.
    if (Application.ActiveSheet.Protection.AllowDeletingRows == false) {
        Application.ActiveSheet.Protect(null, null, null, null, null, null, null, null, null, null, null, null, true)
    }
    console.log("Row 1 can be deleted on this protected worksheet.")
}
```

```javascript
/*本示例显示是否允许删除受保护工作表上的行。*/
function test() {
    console.log(Worksheets.Item(1).Protection.AllowDeletingRows)
}
```
