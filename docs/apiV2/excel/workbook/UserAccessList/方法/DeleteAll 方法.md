---
title: "DeleteAll 方法"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/UserAccessList/%E6%96%B9%E6%B3%95/DeleteAll%20%E6%96%B9%E6%B3%95.html
---
# DeleteAll (方法)

删除有权访问工作表上受保护的区域的所有用户。

## 示例

```javascript
//在本示例中，ET 将删除有权访问活动工作表中第一个受保护的区域的所有用户。本示例假定工作表不受保护。
function test()
{
    let wksSheet = Application.ActiveSheet

    //Remove all users with access to the first protected range.
    wksSheet.Protection.AllowEditRanges.Item(1).Users.DeleteAll()
}
```
