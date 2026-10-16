---
title: "VerticalAlignment (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Range/%E5%B1%9E%E6%80%A7/VerticalAlignment%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: B2EZHOo_
---

# VerticalAlignment (属性)

返回或设置一个 **Variant** 值，它代表指定对象的垂直对齐方式。

## 说明

此属性的值可设为以下常量之一：**xlBottom**、**xlCenter**、**xlDistributed**、**xlJustify**、**xlTop**

## 示例

```javascript
/*本示例将 A1 单元格的垂直对齐方式设置为下对齐*/
function test() {
    Range("A1").VerticalAlignment = xlBottom
}
```

```javascript
/*本示例将 C2 单元格的垂直对齐方式设置为居中*/
function test() {
    Range("C2").VerticalAlignment = xlCenter
}
```
