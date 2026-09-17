---
title: "AllowEditRange 对象"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/AllowEditRange/AllowEditRange%20%E5%AF%B9%E8%B1%A1.html
---
# AllowEditRange (对象)

代表受保护的工作表上可进行编辑的单元格。

## 说明

使用 **[AllowEditRanges](./../AllowEditRanges/AllowEditRanges 对象.html)** 集合的 **Add** 方法或 **Item** 属性可返回 **[AllowEditRange](./../AllowEditRange/AllowEditRange 对象.html)** 对象。

返回 **[AllowEditRange](./../AllowEditRange/AllowEditRange 对象.html)** 对象后，可使用 **ChangePassword** 方法更改密码以访问可在受保护的工作表上编辑的区域。

## 示例

```javascript
/*在此示例中，ET 允许编辑活动工作表上的 A1:A4 范围，通知用户更改此指定区域的密码，然后通知用户更改成功。在运行此代码之前，工作表必须未受到保护。*/
function test() {
    let sheet = Application.ActiveSheet
    sheet.Unprotect()
    let wksPassword = "Enter password for the worksheet"

    //Establish a range that can allow edits on the protected worksheet.
    sheet.Protection.AllowEditRanges.Add("Classified", Range("A1:A4"), wksPassword)
    console.log("Cells A1 to A4 can be edited on the protected worksheet.")

    //Change the password.
    wksPassword = "Enter the new password for the worksheet"
    sheet.Protection.AllowEditRanges("Classified").ChangePassword(wksPassword)

    console.log("The password for these cells has been changed.")
}
```

```javascript
/*本示例将活动工作表上第一个可编辑的单元格区域删除。*/
function test() {
    ActiveSheet.Protection.AllowEditRanges.Item(1).Delete()
}
```
