---
title: "Locked 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/CellFormat/%E5%B1%9E%E6%80%A7/Locked%20%E5%B1%9E%E6%80%A7.html
---
# Locked (属性)

返回或设置一个 **Variant** 值，它指明对象是否已被锁定。

## 说明

如果对象已被锁定，此属性将返回 **True**；如果在工作表处于受保护状态时仍能修改对象，则返回 **False**；如果指定区域既包含锁定单元格又包含不锁定单元格，则返回 **Null**。

## 示例

```javascript
/*此示例设置替换条件中单元格为锁定状态。*/
function test() {
    Application.ReplaceFormat.Locked = true
}
```

```javascript
/*此示例显示查找条件中单元格是否设置为已被锁定状态。*/
function test() {
    console.log(Application.FindFormat.Locked)
}
```
