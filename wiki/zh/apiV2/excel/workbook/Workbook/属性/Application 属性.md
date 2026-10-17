---
title: "Application (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Workbook/%E5%B1%9E%E6%80%A7/Application%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: CnUdLcxA
---

# Application (属性)

如果不使用对象识别符，则该属性返回一个 Application 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 **[Application](./../../Application/Application 对象.html)** 对象。只读。

## 示例

```javascript
/*本示例显示一条有关当前应用程序的消息。*/
function test() {
    if (Application.ActiveWorkbook.Application.Value == "ET") {
        console.log("This is an ET Application object.")
    } else {
        console.log("This is not an ET Application object.")
    }
}
```
