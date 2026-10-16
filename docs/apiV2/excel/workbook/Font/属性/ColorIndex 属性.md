---
title: "ColorIndex 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Font/%E5%B1%9E%E6%80%A7/ColorIndex%20%E5%B1%9E%E6%80%A7.html
---
# ColorIndex (属性)

返回或设置一个 **Variant** 值，它代表字体的颜色。

## 说明

颜色可指定为当前调色板中颜色的索引值，也可指定为下列 **XlColorIndex** 常量之一：

- **xlColorIndexAutomatic**
- **xlColorIndexNone**

## 示例

```javascript
/*本示例将 Sheet1 上 A1 单元格的字体颜色更改为红色*/
function test() {
    Application.Worksheets.Item("Sheet1").Range("A1").Font.ColorIndex = 3
}
```

```javascript
/*此示例判断D3单元格的字体颜色是否为自动配色，不是则设置字体颜色为绿色*/
function test() {
    let font = Application.Worksheets.Item("Sheet1").Range("D3").Font
    if (font.ColorIndex == xlColorIndexAutomatic) {
        console.log(`字体颜色为自动配色`)
    } else {
        font.ColorIndex = 4
    }
}
```
