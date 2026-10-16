---
title: "WriteReservedBy 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Workbook/%E5%B1%9E%E6%80%A7/WriteReservedBy%20%E5%B1%9E%E6%80%A7.html
---
# WriteReservedBy (属性)

返回当前对指定工作簿有写权限的用户的名称。**String** 类型，只读。

## 示例

```javascript
/*如果当前工作簿是写保护的，本示例将显示一则消息，其中包含将该工作簿以写保护方式保存的用户的名称。*/
function test() {
    if (ActiveWorkbook.WriteReserved) {
        console.log(`Please contact ${ActiveWorkbook.WriteReservedBy} if you need to insert data in this workbook.`)
    }
}
```

```javascript
/*本示例设置工作簿的写密码，并通知用户有写权限的用户的名称。*/
function test() {
    Workbooks.Item(1).WritePassword = "Y6dh!et5"
    console.log(Workbooks.Item(1).WriteReservedBy)
}
```
