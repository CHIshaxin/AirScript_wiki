---
title: "InCellDropdown (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Validation/%E5%B1%9E%E6%80%A7/InCellDropdown%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: BE9u77c3
---

# InCellDropdown (属性)

如果数据有效性显示含有有效取值的下拉列表，则该属性值为 **True**。**Boolean** 类型，可读写。

## 说明

如果数据有效性检验类型不是 **xlValidateList**，将忽略该属性。

可用 **[Validation](./../../Validation/Validation 对象.html)** 对象的 **Add** 方法或 **Modify** 方法的 *Minimum* 参数来指定包含有效数据的区域。

## 示例

```javascript
/*本示例向单元格 E5 添加数据有效性检验。单元格区域 A1:A10 包含单元格的有效取值，且该单元格显示含有这些有效取值的下拉列表。*/
function test() {
    let validation = Application.Range("E5").Validation
    validation.Add(xlValidateList, xlValidAlertStop, xlBetween, "=$A$1:$A$10")
    validation.InCellDropdown = true
}
```

```javascript
/*本示例显示活动工作表上单元格 E1 的数据有效性检验是否展示含有有效取值的下拉列表。*/
function test() {
    let validation = ActiveSheet.Range("E1").Validation
    console.log(validation.InCellDropdown)
}
```
