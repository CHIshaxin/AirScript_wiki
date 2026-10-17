---
title: "Type (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Validation/%E5%B1%9E%E6%80%A7/Type%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: DEMF3Xf3
---

# Type (属性)

返回一个包含 **[XlDVType](./../../枚举/XlDVType 枚举.html)** 常量的 **Long** 值，它代表区域的数据类型有效性验证。

## 示例

```javascript
/*本示例显示第一张工作表上单元格 B1 的数据有效性验证的类型是否为xlValidateWholeNumber。*/
function test() {
    let validation = Application.Worksheets.Item(1).Range("B1").Validation
    console.log(validation.Type == xlValidateWholeNumber)
}
```

```javascript
/*本示例判断活动工作表上 B1 单元格的数据有效性验证的类型是否为xlValidateDate，并通知用户。*/
function test() {
    let validation = Application.ActiveSheet.Range("B1").Validation
    if (validation.Type == xlValidateDate) {
        console.log("数据有效性验证的类型为xlValidateDate")
    } else {
        console.log("数据有效性验证的类型不为xlValidateDate")
    }
}
```
