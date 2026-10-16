---
title: "AllowInsertingRows (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Protection/%E5%B1%9E%E6%80%A7/AllowInsertingRows%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: CSq6dWXW
---

# AllowInsertingRows (属性)

如果允许用户在受保护的工作表上插入行，则返回 **True**。**Boolean** 类型，只读。

## 说明

可以使用 **Protect** 方法参数设置 **AllowInsertingRows** 属性。

## 示例

```javascript
/*本示例允许用户在受保护的工作表上插入行，并通知用户。*/
function test() {
    Application.ActiveSheet.Unprotect()
    // Allow rows to be inserted on a protected worksheet.
    if (Application.ActiveSheet.Protection.AllowInsertingRows == false) {
        Application.ActiveSheet.Protect(null, null, null, null, null, null, null, null, null, true)
    }
    console.log("Rows can be inserted on this protected worksheet.")
}
```

```javascript
/*本示例显示是否允许在受保护的第二张工作表上插入行。*/
function test() {
    console.log(Worksheets.Item(2).Protection.AllowInsertingRows)
}
```
