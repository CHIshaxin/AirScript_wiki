---
title: "Operator (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Validation/%E5%B1%9E%E6%80%A7/Operator%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: 5OVUk-xX
---

# Operator (属性)

返回一个 **Long** 值，它代表数据有效性的运算符。

## 示例

```javascript
/*本示例显示单元格 B1 的数据有效性的运算符是否为xlBetween。*/
function test() {
    console.log(Range("B1").Validation.Operator == xlBetween)
}
```

```javascript
/*本示例判断活动工作表上 B1 单元格的数据有效性的运算符是否为xlGreater，并通知用户。*/
function test() {
    let validation = Application.ActiveSheet.Range("B1").Validation
    if (validation.Operator == xlGreater) {
        console.log("数据有效性的运算符为xlGreater")
    } else {
        console.log("数据有效性的运算符不为xlGreater")
    }
}
```
