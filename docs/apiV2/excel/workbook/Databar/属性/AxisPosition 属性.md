---
title: "AxisPosition 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Databar/%E5%B1%9E%E6%80%A7/AxisPosition%20%E5%B1%9E%E6%80%A7.html
---
# AxisPosition (属性)

返回或设置由条件格式规则指定的数据条的坐标轴位置。可读/写。

## 说明

数据条的坐标轴只在以下情况下显示：**AxisPosition** 属性为 **xlDataBarAxisAutomatic** 或 **xlDataBarAxisMidpoint**；以及指定了数据条条件格式规则的值区域中有负值。**AxisPosition** 属性的默认值为 **xlDataBarAxisAutomatic**。

## 返回值

**[XlDataBarAxisPosition](./../../枚举/XlDataBarAxisPosition 枚举.html)**

## 示例

```javascript
/*本示例选择一个单元格区域，添加数据条格式，然后将坐标轴位置设置为当存在负值时显示在单元格当中。*/
function test() {
    Range("A1:A10").Select()
    Range("A1:A10").Activate()
    let databar = Selection.FormatConditions.AddDatabar()
    databar.AxisPosition = xlDataBarAxisMidpoint
}
```

```javascript
/*本示例显示工作表 Sheet1 上区域 A1:A10 中第一个为数据条的条件格式的坐标轴位置是否为xlDataBarAxisNone。*/
function test() {
    let databar = Application.Worksheets.Item("Sheet1").Range("A1:A10").FormatConditions.Item(1)
    console.log(databar.AxisPosition == xlDataBarAxisNone)
}
```
