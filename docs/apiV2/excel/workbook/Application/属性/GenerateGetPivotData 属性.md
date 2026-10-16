---
title: "GenerateGetPivotData 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Application/%E5%B1%9E%E6%80%A7/GenerateGetPivotData%20%E5%B1%9E%E6%80%A7.html
---
# GenerateGetPivotData (属性)

当 ET 能够获取数据透视表数据时，返回 **True**。**Boolean** 类型，可读写。

## 示例

```javascript
/*本示例为 ET 确定获取数据透视表数据的状态，并通知用户。本示例假定数据透视表位于活动工作表上。*/
function test() {
    // Determine the ability to get PivotTable report data and notify user.
    if (Application.GenerateGetPivotData) {
        console.log("The ability to get PivotTable report data is enabled.")
    } else {
        console.log("The ability to get PivotTable report data is disabled.")
    }
}
```

```javascript
/*本示例将 ET 设置为不能获取数据透视表数据。*/
function test() {
    Application.GenerateGetPivotData = false
}
```
