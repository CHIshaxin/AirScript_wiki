---
title: "ChangePassword (方法)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/AllowEditRange/%E6%96%B9%E6%B3%95/ChangePassword%20%E6%96%B9%E6%B3%95.html
updated: 1789700916000
hash: BZ7AW46G
---

# ChangePassword (方法)

更改受保护的工作表中可以进行编辑的区域的密码。

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
|---|---|---|---|
| Password | string | 必填 | 新密码。 |

## 示例

```javascript
/*在此示例中，ET 允许编辑活动工作表上的 A1:A4 范围，通知用户更改此指定区域的密码，然后通知用户更改成功。在运行此代码之前，工作表必须未受到保护。*/
function test() {
    let sheet = Application.ActiveSheet
    sheet.Unprotect()

    //Establish a range that can allow edits on the protected worksheet.
    let strPassword = "Please enter the password for the range"
    sheet.Protection.AllowEditRanges.Add("Classified", Range("A1:A4"), strPassword)

    strPassword = "Please enter the new password for the range"

    //Change the password.
    sheet.Protection.AllowEditRanges("Classified").ChangePassword(strPassword)

    console.log("The password for these cells has been changed.")
}
```

```javascript
/*本示例将活动工作表上第一个可以进行编辑的区域的密码更改为“Y6dh!et5”。*/
function test() {
    ActiveSheet.Protection.AllowEditRanges.Item(1).ChangePassword("Y6dh!et5")
}
```
