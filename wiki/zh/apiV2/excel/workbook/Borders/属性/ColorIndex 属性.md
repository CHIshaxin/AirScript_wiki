---
title: "ColorIndex (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Borders/%E5%B1%9E%E6%80%A7/ColorIndex%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: CH1wHB8a
---

# ColorIndex (属性)

返回或设置一个 **Variant** 值，它代表全部四条边框的颜色。

## 说明

如果全部四条边框不是同一种颜色，此属性返回 **Null**。

颜色可指定为当前调色板中颜色的索引值，也可指定为下列 **XlColorIndex** 常量之一：

- **xlColorIndexAutomatic**
- **xlColorIndexNone**

## 示例

```javascript
/*本示例显示 A2 单元格的边框颜色是否为无颜色。*/
function test() {
    console.log(Range("A2").Borders.ColorIndex == xlColorIndexNone)
}
```

```javascript
/*本示例将第一个工作表中 E3 单元格的边框颜色设置为黄色。*/
function test() {
    Worksheets.Item(1).Range("E3").Borders.ColorIndex = 6
}
```
