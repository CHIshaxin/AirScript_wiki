---
title: "Formula1 (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Validation/%E5%B1%9E%E6%80%A7/Formula1%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: CgLcBXCs
---

# Formula1 (属性)

返回与条件格式或者数据有效性相关联的值或表达式。可为常量值、字符串值、单元格引用或公式。**String**型，只读。

## 示例

```javascript
/*本示例显示单元格 B1 的数据有效性验证的第一部分相关联的表达式。*/
function test() {
    let validation = ActiveSheet.Range("B1").Validation
    console.log(validation.Formula1)
}
```

```javascript
/*本示例判断如果活动工作表上单元格 B1 的数据有效性验证的第一部分相关联的表达式为“=A1”，则删除该数据有效性验证。*/
function test() {
    let validation = ActiveSheet.Range("B1").Validation
    if (validation.Formula1 == "=A1") {
        validation.Delete()
    }
}
```
