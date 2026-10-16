---
title: "CellDragAndDrop (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Application/%E5%B1%9E%E6%80%A7/CellDragAndDrop%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: gscUupw5
---

# CellDragAndDrop (属性)

如果启用单元格拖放功能，则该属性值为 **True**。**Boolean** 类型，可读写。

## 示例

```javascript
/*本示例启用单元格拖放功能。*/
function test() {
    Application.CellDragAndDrop = true
}
```

```javascript
/*本示例演示CellDragAndDrop属性的数据类型，并将CellDragAndDrop属性的值设置为false。*/
function test() {
    console.log(typeof Application.CellDragAndDrop)
    Application.CellDragAndDrop = false
}
```
