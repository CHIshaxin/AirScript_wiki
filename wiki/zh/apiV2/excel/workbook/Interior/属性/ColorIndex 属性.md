---
title: "ColorIndex (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Interior/%E5%B1%9E%E6%80%A7/ColorIndex%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: CUyDvqFd
---

# ColorIndex (属性)

返回或设置一个**Variant** 值，它代表内部颜色。

## 说明

颜色可指定为当前调色板中颜色的索引值，也可指定为下列 **XlColorIndex** 常量之一：

- **xlColorIndexAutomatic**
- **xlColorIndexNone**

## 示例

```javascript
/*此示例将 Sheet1 中 C2 单元格内部设置为红色。*/
function test() {
    Application.Worksheets("Sheet1").Range("C2").Interior.ColorIndex = 3
}
```

```javascript
/*此示例将活动工作表中 A1:D5 区域单元格内部设置为无颜色。*/
function test() {
    ActiveSheet.Range("A1:D5").Interior.ColorIndex = xlColorIndexNone
}
```
