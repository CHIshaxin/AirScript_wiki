---
title: "Unprotect (方法)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/AllowEditRange/%E6%96%B9%E6%B3%95/Unprotect%20%E6%96%B9%E6%B3%95.html
updated: 1789700916000
hash: CeJMDYDA
---

# Unprotect (方法)

取消可编辑的单元格区域的保护。如果可编辑的单元格区域不是受保护的，则此方法不起作用。

## 说明

如果您忘记了密码，将不能取消可编辑的单元格区域的保护。建议将密码和对应文档名妥善保存。

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
|---|---|---|---|
| Password | any | 可选 | 一个字符串，它指定用于解除单元格区域保护的密码，此密码是区分大小写的。如果单元格区域不设密码保护，则忽略此参数。 |

## 示例

```javascript
/*本示例取消第一张工作表上第一个可编辑的单元格区域的保护。*/
function test() {
    Worksheets.Item(1).Protection.AllowEditRanges.Item(1).Unprotect()
}
```

```javascript
/*本示例取消活动工作表上第二个可编辑的单元格区域的保护。*/
function test() {
    ActiveSheet.Protection.AllowEditRanges.Item(2).Unprotect("Y6dh!et5")
}
```
