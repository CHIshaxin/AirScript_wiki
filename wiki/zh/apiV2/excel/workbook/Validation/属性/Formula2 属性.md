---
title: "Formula2 (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Validation/%E5%B1%9E%E6%80%A7/Formula2%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: Cts4sao5
---

# Formula2 (属性)

返回与条件格式或数据有效性验证第二部分相关联的值或表达式。仅用于数据有效性条件格式 **Operator** 属性为 **xlBetween** 或 **xlNotBetween**的情况。可为常量值、字符串值、单元格引用或公式。**String** 类型，只读。

## 示例

```javascript
/*本示例显示单元格 B2 的数据有效性验证的第二部分相关联的值。*/
function test() {
    console.log(Range("B2").Validation.Formula2)
}
```

```javascript
/*本示例将活动工作表上单元格 B2 的数据有效性验证的第二部分相关联的值赋值给 C1 单元格。*/
function test() {
    let validation = ActiveSheet.Range("B2").Validation
    Range("C1").Value2 = validation.Formula2
}
```
