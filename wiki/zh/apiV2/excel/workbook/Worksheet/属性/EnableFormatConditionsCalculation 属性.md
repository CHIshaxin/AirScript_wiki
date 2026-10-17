---
title: "EnableFormatConditionsCalculation (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Worksheet/%E5%B1%9E%E6%80%A7/EnableFormatConditionsCalculation%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: KZbzat-W
---

# EnableFormatConditionsCalculation (属性)

返回或设置是否在需要时自动计算条件格式。可读/写 **Boolean** 类型。

## 说明

如果设置为 True（默认值），则在需要时自动计算条件格式。如果设置为 False，则不重新计算条件格式。任何先前应用的条件格式仍可见，但不会在单元格值或“应用于”区域更改时进行更新。

使用此标志，是为了使 JavaScript 程序员可以在计算规则前进行完全配置。在条件应用于较大区域，因而性能较低的情况下，这种方式尤其有用。

## 示例

```javascript
/*本示例判断如果条件格式不会重新计算，则开启重新计算条件格式。*/
function test() {
    if (!ActiveSheet.EnableFormatConditionsCalculation) {
        ActiveSheet.EnableFormatConditionsCalculation = true
    }
}
```

```javascript
/*本示例判断工作表 Sheet1 是否自动计算条件格式，并通知用户。*/
function test() {
    if (Application.Sheets.Item("Sheet1").EnableFormatConditionsCalculation) {
        console.log("工作表 Sheet1 自动计算条件格式")
    } else {
        console.log("工作表 Sheet1 不会自动计算条件格式")
    }
}
```
