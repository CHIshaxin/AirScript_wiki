---
title: "MaxChange (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Application/%E5%B1%9E%E6%80%A7/MaxChange%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: CyGXtoCe
---

# MaxChange (属性)

返回或设置 ET 处理循环引用时迭代之间的最大变化值。**Double** 类型，可读写。

## 说明

**MaxIterations** 属性设置 ET 在解决循环引用时可使用的最多迭代次数。

## 示例

```javascript
/*本示例将每次迭代的最大变化值设为0.1。*/
function test() {
    Application.MaxChange = 0.1
}
```

```javascript
/*本示例演示将MaxChange属性的值设置为0.5，并显示该属性的值。*/
function test() {
    Application.MaxChange = 0.5
    console.log("MaxChange属性的值为：" + Application.MaxChange)
}
```
