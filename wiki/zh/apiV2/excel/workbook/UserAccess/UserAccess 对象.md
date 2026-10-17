---
title: "UserAccess (对象)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/UserAccess/UserAccess%20%E5%AF%B9%E8%B1%A1.html
updated: 1789700916000
hash: B08wPgrj
---

# UserAccess (对象)

代表对受保护区域的用户访问。

## 说明

使用 UserAccessList 集合的 **Add**方法或 Item 属性可返回一个 **[UserAccess](./../UserAccess/UserAccess 对象.html)** 对象。

一旦返回了 **[UserAccess](./../UserAccess/UserAccess 对象.html)** 对象，您就可以使用 **AllowEdit**属性来确定是否允许访问工作表中某个特定区域。下例添加一个在受保护的工作表上可编辑的区域，并通知用户该区域的标题。

```javascript
function test(){
    let wksSheet = Application.ActiveSheet

    //Add a range that can be edited on the protected worksheet.
    wksSheet.Protection.AllowEditRanges.Add("Test", Range("A1"))

    //Notify the user the title of the range that can be edited.
    console.log(wksSheet.Protection.AllowEditRanges.Item(1).Title)
}
```
