---
title: "AllowInsertingColumns 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Protection/%E5%B1%9E%E6%80%A7/AllowInsertingColumns%20%E5%B1%9E%E6%80%A7.html
---
# AllowInsertingColumns (属性)

如果允许在受保护的工作表上插入列，则返回 **True**。**Boolean** 类型，只读。

## 说明

默认情况下，插入的列继承了其左边的列的格式设置，这意味着该列可能包含锁定的单元格。也就是说，用户可能不能删除插入的列。

可以使用 **Protect** 方法参数设置 **AllowInsertingColumns** 属性。

## 示例

```javascript
/*本示例允许用户在受保护的工作表上插入列，并通知用户。*/
function test() {
    Application.ActiveSheet.Unprotect()

    // Allow columns to be inserted on a protected worksheet.
    if (Application.ActiveSheet.Protection.AllowInsertingColumns == false) {
        Application.ActiveSheet.Protect(null, null, null, null, null, null, null, null, true)
    }
    console.log("Columns can be inserted on this protected worksheet.")
}
```

```javascript
/*本示例显示是否允许在受保护的第二张工作表上插入列。*/
function test() {
    console.log(Worksheets.Item(2).Protection.AllowInsertingColumns)
}
```
