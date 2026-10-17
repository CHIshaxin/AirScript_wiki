---
title: "IgnoreBlank (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Validation/%E5%B1%9E%E6%80%A7/IgnoreBlank%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: DB-JChsC
---

# IgnoreBlank (属性)

如果指定区域内的数据有效性检验允许空值，则该值为 **True**。**Boolean** 类型，可读写。

## 说明

当 **IgnoreBlank** 属性为 **True** 时，如果单元格为空白，或者 **MinVal** 或 **MaxVal** 属性所引用的单元格为空白，仍认为该单元格数据有效。

## 示例

```javascript
/*本示例将单元格 E5 的数据有效性检验设置为允许有空值。*/
function test() {
    Application.Range("E5").Validation.IgnoreBlank = true
}
```

```javascript
/*本示例显示活动工作表上单元格 E5 的数据有效性检验是否允许有空值。*/
function test() {
    let validation = ActiveSheet.Range("E5").Validation
    console.log(validation.IgnoreBlank)
}
```
