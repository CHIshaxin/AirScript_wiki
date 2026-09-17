---
title: "UserAccessList 对象"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/UserAccessList/UserAccessList%20%E5%AF%B9%E8%B1%A1.html
---
# UserAccessList (对象)

代表受保护区域用户访问权限的 **[UserAccess](./../UserAccess/UserAccess 对象.html)** 对象的集合。

## 说明

使用受保护的 **[Range](./../Range/Range 对象.html)** 对象的 **Users** 属性可返回一个 **[UserAccessList](./../UserAccessList/UserAccessList 对象.html)** 集合。

一旦返回了 **[UserAccessList](./../UserAccessList/UserAccessList 对象.html)** 集合，您就可以使用 **Count** 属性来确定能够访问受保护区域的用户的数量。在下例中，ET 通知用户能够访问第一个受保护区域的用户的数量。本示例假定活动工作表中存在受保护区域。

```javascript
function test(){
    let wksSheet = Application.ActiveSheet

    //Notify the user the number of users that can access the protected range.
    console.log(wksSheet.Protection.AllowEditRanges.Item(1).Users.Count)
}
```
