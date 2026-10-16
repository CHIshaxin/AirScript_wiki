---
title: "AllowInsertingHyperlinks (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Protection/%E5%B1%9E%E6%80%A7/AllowInsertingHyperlinks%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: CFEIipwP
---

# AllowInsertingHyperlinks (属性)

如果允许在受保护的工作表上插入超链接，则返回 **True**。**Boolean** 类型，只读。

## 说明

在受保护的工作表上，只能将超链接插入未锁定或未保护的单元格中。

可以使用 **Protect** 方法参数设置 **AllowInsertingHyperlinks** 属性。

## 示例

```javascript
/*本示例允许用户在受保护的工作表上的单元格 A1 中插入超链接，并通知用户。*/
function test() {
    Application.ActiveSheet.Unprotect()
    // Unlock cell A1.
    Range("A1").Locked = false
    // Allow hyperlinks to be inserted on a protected worksheet.
    if (Application.ActiveSheet.Protection.AllowInsertingHyperlinks == false) {
        Application.ActiveSheet.Protect(null, null, null, null, null, null, null, null, null, null, true)
    }
    console.log("Hyperlinks can be inserted on this protected worksheet.")
}
```

```javascript
/*本示例显示是否允许在受保护的第二张工作表上插入超链接。*/
function test() {
    console.log(Worksheets.Item(2).Protection.AllowInsertingHyperlinks)
}
```
