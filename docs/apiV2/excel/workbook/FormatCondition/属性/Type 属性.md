---
title: "Type 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/FormatCondition/%E5%B1%9E%E6%80%A7/Type%20%E5%B1%9E%E6%80%A7.html
---
# Type (属性)

返回一个包含 **xlFormatConditionType** 值的 **Long** 值，它代表对象类型。

## 示例

```javascript
/*本示例显示活动工作表上区域 A1:A10 中第二个条件格式的类型是否为xlColorScale。*/
function test() {
    let formatCondition = ActiveSheet.Range("A1:A10").FormatConditions.Item(2)
    console.log(formatCondition.Type == xlColorScale)
}
```

```javascript
/*本示例判断活动工作表上区域 A1:A10 中第一个条件格式的类型是否为DataBar，并通知用户。*/
function test() {
    let formatCondition = ActiveSheet.Range("A1:A10").FormatConditions.Item(1)
    if (formatCondition.Type == xlDatabar) {
        console.log("类型为 DataBar")
    } else {
        console.log("类型不为 DataBar")
    }
}
```
