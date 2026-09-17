---
title: "AllowFormattingRows 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Protection/%E5%B1%9E%E6%80%A7/AllowFormattingRows%20%E5%B1%9E%E6%80%A7.html
---
# AllowFormattingRows (属性)

如果允许对受保护的工作表上的行进行格式设置，则返回 **True**。**Boolean** 类型，只读。

## 说明

可以使用 **Protect** 方法参数设置 **AllowFormattingRows** 属性。

## 示例

```javascript
/*本示例允许用户对受保护的工作表上的行进行格式设置，并通知用户。*/
function test() {
    Application.ActiveSheet.Unprotect()
    // Allow rows to be formatted on a protected worksheet.
    if (Application.ActiveSheet.Protection.AllowFormattingRows == false) {
        Application.ActiveSheet.Protect(null, null, null, null, null, null, null, true)
    }
    console.log("Rows can be formatted on this protected worksheet.")
}
```

```javascript
/*本示例显示是否允许对受保护的工作表上的行进行格式设置。*/
function test() {
    console.log(Worksheets.Item(1).Protection.AllowFormattingRows)
}
```
