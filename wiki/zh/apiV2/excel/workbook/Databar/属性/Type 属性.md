---
title: "Type (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Databar/%E5%B1%9E%E6%80%A7/Type%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: OsJqMYQd
---

# Type (属性)

返回 **[XlFormatConditionType](./../../枚举/XlFormatConditionType 枚举.html)** 枚举的常量之一，该常量指定条件格式的类型。只读。

## 说明

此属性将始终返回 **Long** 值“4”，等价于 **xlDatabar**。

## 示例

```javascript
/*本示例显示活动工作表上区域 A1:A10 中第一个条件格式（数据条）的类型是否为xlColorScale。*/
function test() {
    let databar = ActiveSheet.Range("A1:A10").FormatConditions.Item(1)
    console.log(databar.Type == xlColorScale)
}
```

```javascript
/*本示例判断活动工作表上区域 A1:A10 中第一个条件格式（数据条）的类型是否为DataBar，并通知用户。*/
function test() {
    let databar = ActiveSheet.Range("A1:A10").FormatConditions.Item(1)
    if (databar.Type == xlDatabar) {
        console.log("类型为 DataBar")
    } else {
        console.log("类型不为 DataBar")
    }
}
```
