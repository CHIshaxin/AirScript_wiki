---
title: "ProductCode (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Application/%E5%B1%9E%E6%80%A7/ProductCode%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: CFQiEgWi
---

# ProductCode (属性)

返回 ET 的全球唯一标识符 (GUID)。**String** 类型，只读。

## 示例

```javascript
/*此示例显示 ET 的 GUID。*/
function test() {
    console.log(Application.ProductCode)
}
```

```javascript
/*本示例显示 ProductCode 属性的数据类型。*/
function test() {
    console.log("ProductCode属性的数据类型为：" + typeof Application.ProductCode)
}
```
