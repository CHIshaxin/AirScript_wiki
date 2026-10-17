---
title: "EnableCalculation (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Worksheet/%E5%B1%9E%E6%80%A7/EnableCalculation%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: B_uq9V3N
---

# EnableCalculation (属性)

如果在必要的情况下 ET 自动重新计算工作表，则该值为 **True**。如果 ET 不重新计算工作表，则该值为 **False**。**Boolean** 类型，可读写。

## 说明

当本属性的值为 **False** 时，则不能请求重新计算。如果将其值从 **False** 更改为 **True**，则 ET 将重新计算工作表。

## 示例

```javascript
/*本示例将 WPS表格 设置为不自动重新计算第一个工作表。*/
function test() {
    Application.Worksheets.Item(1).EnableCalculation = false
}
```

```javascript
/*本示例判断活动工作表是否开启自动重新计算，并通知用户。*/
function test() {
    if (Application.ActiveSheet.EnableCalculation) {
        console.log("活动工作表开启自动重新计算")
    } else {
        console.log("活动工作表未开启自动重新计算")
    }
}
```
