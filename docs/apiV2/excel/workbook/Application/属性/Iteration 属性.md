---
title: "Iteration 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Application/%E5%B1%9E%E6%80%A7/Iteration%20%E5%B1%9E%E6%80%A7.html
---
# Iteration (属性)

如果 ET 使用迭代来处理循环引用，则该属性值为 **True**。**Boolean** 类型，可读写。

## 示例

```javascript
/*本示例将 Iteration 属性设置为true，以使用迭代来处理循环引用。*/
function test() {
    Application.Iteration = true
}
```

```javascript
/*本示例演示 Iteration 属性的数据类型，并将 Iteration 属性的值设置为false。*/
function test() {
    console.log("Iteration属性的数据类型为：" + typeof Application.Iteration)
    Application.Iteration = false
}
```
