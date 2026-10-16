---
title: "BarFillType (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Databar/%E5%B1%9E%E6%80%A7/BarFillType%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: DhWltVfQ
---

# BarFillType (属性)

返回或设置数据条如何填充颜色。可读/写。

## 说明

**BarFillType** 属性的默认值为 **xlDataBarFillGradient**。

## 返回值

**[XlDataBarFillType](./../../枚举/XlDataBarFillType 枚举.html)**

## 示例

```javascript
/*本示例选择一个单元格区域，将数据条条件格式规则添加到该区域，然后将数据条的填充色设置为纯色。*/
function test() {
    Range("A1:A10").Select()
    Range("A1:A10").Activate()
    let databar = Selection.FormatConditions.AddDatabar()
    databar.BarFillType = xlDataBarFillSolid
}
```

```javascript
/*本示例显示活动工作表上区域 A1:A10 中第一个为数据条的条件格式的填充色是否为渐变色。*/
function test() {
    let databar = ActiveSheet.Range("A1:A10").FormatConditions.Item(1)
    console.log(databar.BarFillType == xlDataBarFillGradient)
}
```
