---
title: "ExtendList (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Application/%E5%B1%9E%E6%80%A7/ExtendList%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: BjpxJPNa
---

# ExtendList (属性)

如果 ET 自动将格式和公式扩展到清单中新增的数据上，则该属性值为 **True**。**Boolean** 类型，可读写。

## 说明

要实现格式和公式的扩展，则清单中新行或新列前五行或五列中至少应有三行或三列必须具有这些格式和公式，且数据必须添加到清单的底部或右方。

## 示例

```javascript
/*本示例设置 ET ，使其不对以后添加到清单中的数据应用格式和公式。*/
function test() {
    Application.ExtendList = false
}
```

```javascript
/*本示例演示 ExtendList 属性的数据类型，并将 ExtendList 属性的值设置为true。*/
function test() {
    console.log("ExtendList属性的数据类型为：" + typeof Application.ExtendList)
    Application.ExtendList = true
}
```
