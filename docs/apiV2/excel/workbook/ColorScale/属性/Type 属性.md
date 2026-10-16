---
title: "Type 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/ColorScale/%E5%B1%9E%E6%80%A7/Type%20%E5%B1%9E%E6%80%A7.html
---
# Type (属性)

返回 **[XlFormatConditionType](./../../枚举/XlFormatConditionType 枚举.html)** 枚举的常量之一，该常量指定条件格式的类型。只读。

## 说明

此属性将始终返回 **Long** 值“3”，这相当于 **xlColorScale**。

## 示例

```javascript
/*本示例显示活动工作表上区域 C1:C10 中第一个条件格式的类型是否为色阶条件格式。*/
function test() {
    let colorScale = ActiveSheet.Range("C1:C10").FormatConditions.Item(1)
    console.log(colorScale.Type == xlColorScale)
}
```

```javascript
/*本示例判断第一张工作表上区域 C1:C10 中第二个条件格式的类型是否为色阶条件格式，并通知用户。*/
function test() {
    let colorScale = Worksheets.Item(1).Range("C1:C10").FormatConditions.Item(2)
    if (colorScale.Type == xlColorScale) {
        console.log("类型为色阶条件格式")
    } else {
        console.log("类型不为色阶条件格式")
    }
}
```
