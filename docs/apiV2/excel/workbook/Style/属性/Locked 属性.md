---
title: "Locked 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Style/%E5%B1%9E%E6%80%A7/Locked%20%E5%B1%9E%E6%80%A7.html
---
# Locked (属性)

返回或设置一个 **Boolean** 值，它指明对象是否已被锁定。

## 说明

如果对象已被锁定，此属性将返回 **True**；如果在工作表处于受保护状态时仍能修改对象，则返回 **False**。

## 示例

```javascript
/*本示例显示A1单元格的样式是否锁定。*/
function test() {
    let range = Application.Worksheets.Item(1).Range("A1")
    console.log(range.Style.Locked)
}
```

```javascript
/*本示例将C2单元格的样式解除锁定。*/
function test() {
    let range = ActiveSheet.Range("C2")
    range.Style.Locked = false
}
```
