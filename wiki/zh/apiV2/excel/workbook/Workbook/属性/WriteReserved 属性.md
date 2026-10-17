---
title: "WriteReserved (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Workbook/%E5%B1%9E%E6%80%A7/WriteReserved%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: DjYfyMuo
---

# WriteReserved (属性)

如果指定工作簿是写保护的，则该属性的值为 **True**。**Boolean** 类型，只读。

## 说明

可用 **SaveAs** 方法设置该属性。

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
/*本示例另存活动工作簿，判断活动工作簿是否为写保护，并通知用户。*/
function test() {
    ActiveWorkbook.SaveAs(ActiveWorkbook.FullName, undefined, undefined, "Y6dh!et5")
    if (ActiveWorkbook.WriteReserved) {
        console.log("活动工作簿已开启写保护")
    }
}
```
