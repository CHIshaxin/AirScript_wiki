---
title: "Application (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Application/%E5%B1%9E%E6%80%A7/Application%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: DNBT06vl
---

# Application (属性)

该属性返回一个代表 ET 应用程序的 **[Application](./../../Application/Application 对象.html)** 对象。只读。

## 示例

```javascript
/*本示例显示一条有关创建 myObject 的应用程序的消息*/
function test() {
    let myObject = Application
    console.log(myObject.Application.Value)
}
```

```javascript
/*此示例演示将应用程序对象的界面隐藏。*/
function test() {
    Application.Application.Visible = false
}
```
