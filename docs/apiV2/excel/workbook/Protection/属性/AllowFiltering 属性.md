---
title: "AllowFiltering 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Protection/%E5%B1%9E%E6%80%A7/AllowFiltering%20%E5%B1%9E%E6%80%A7.html
---
# AllowFiltering (属性)

如果允许用户使用工作表受保护之前设置的“自动筛选”，则返回 **True**。**Boolean** 类型，只读。

## 说明

可以使用 **Protect** 方法参数设置 **AllowFiltering** 属性。

**AllowFiltering** 属性允许用户更改已有的“自动筛选”上的筛选条件。用户不能创建或删除受保护的工作表上的“自动筛选”。

若要筛选受保护的工作表上的单元格，则必须先取消对这些单元格的锁定。

## 示例

```javascript
/*本示例允许用户筛选受保护的工作表上的第 1 行并通知用户。*/
function test() {
    Application.ActiveSheet.Unprotect()
    // Unlock row 1.
    Rows.Item("1:1").Locked = false
    // Allow row 1 to be filtered on a protected worksheet.
    if (Application.ActiveSheet.Protection.AllowFiltering == false) {
        Application.ActiveSheet.Protect(null, null, null, null, null, null, null, null, null, null, null, null, null, null, true)
    }
    console.log("Row 1 can be filtered on this protected worksheet.")
}
```

```javascript
/*本示例显示是否允许用户使用工作表受保护之前设置的“自动筛选”。*/
function test() {
    console.log(Worksheets.Item(1).Protection.AllowFiltering)
}
```
