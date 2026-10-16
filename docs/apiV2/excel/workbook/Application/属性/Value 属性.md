---
title: "Value 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Application/%E5%B1%9E%E6%80%A7/Value%20%E5%B1%9E%E6%80%A7.html
---
# Value (属性)

返回一个代表应用程序名称的 **String** 值。

## 说明

此属性始终返回 "Microsoft Excel"。

## 示例

```javascript
/*此示例显示应用程序名称的长度*/
function test() {
    console.log(Application.Value.length)
}
```

```javascript
/*此示例显示应用程序名称*/
function test() {
    console.log(Application.Value)
}
```
