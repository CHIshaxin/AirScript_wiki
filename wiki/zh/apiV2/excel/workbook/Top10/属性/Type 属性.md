---
title: "Type (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Top10/%E5%B1%9E%E6%80%A7/Type%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: BPYUqvLa
---

# Type (属性)

返回 **[XlFormatConditionType](./../../枚举/XlFormatConditionType 枚举.html)** 枚举的常量之一，该常量指定条件格式的类型。只读。

## 说明

此属性将始终返回 **Long** 值“5”，等价于 **xlTop10**。

## 示例

```javascript
/*本示例显示活动工作表上区域 A1:A10 中第二个条件格式的类型是否为xlTop10。*/
function test() {
    let top = ActiveSheet.Range("A1:A10").FormatConditions.Item(2)
    console.log(top.Type == xlTop10)
}
```

```javascript
/*本示例判断第一张工作表上区域 A1:A10 中第一个条件格式的类型是否为xlTop10，并通知用户。*/
function test() {
    let formatCondition = Worksheets.Item(1).Range("A1:A10").FormatConditions.Item(1)
    if (formatCondition.Type == xlTop10) {
        console.log("类型为 xlTop10")
    } else {
        console.log("类型不为 xlTop10")
    }
}
```
