---
title: "Font 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Range/%E5%B1%9E%E6%80%A7/Font%20%E5%B1%9E%E6%80%A7.html
---
# Font (属性)

返回一个 **[Font](./../../Font/Font 对象.html)** 对象，它代表指定对象的字体。

## 示例

```javascript
/*此示例判断单元格 A1 的字体名称是否为 Arial，并通知用户。*/
function test() {
    Application.Range("A1").Select()
    // Determine if the font name for selected cell is Arial.
    if (Application.Range("A1").Font.Name == "Arial") {
        console.log("The font name for this cell is 'Arial'")
    } else {
        console.log("The font name for this cell is not 'Arial'")
    }
}
```

```javascript
/*此示例将A1单元格字体加粗*/
function test() {
    Range("A1").Font.Bold = true
}
```
