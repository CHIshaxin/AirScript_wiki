---
title: "Cells 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Range/%E5%B1%9E%E6%80%A7/Cells%20%E5%B1%9E%E6%80%A7.html
---
# Cells (属性)

返回一个 Range 对象，该对象代表指定区域中的单元格。

## 说明

因为 **Item** 属性是 **[Range](./../../Range/Range 对象.html)** 对象的默认属性，所以可以在 **Cells** 关键字后面紧接着指定行和列索引。有关详细信息，请参阅 **Item** 属性和本主题的示例。

在不使用对象识别符的情况下，使用此属性将返回一个 **[Range](./../../Range/Range 对象.html)** 对象，它代表活动工作表中所有的单元格。

## 示例

```javascript
/*本示例将 Sheet1 上单元格区域 A1: C5 的字体样式设置为斜体。*/
function test() {
    Application.Worksheets.Item("Sheet1").Activate()
    Application.Worksheets.Item("Sheet1").Range(Cells.Item(1, 1), Cells.Item(5, 3)).Font.Italic = true
}
```

```javascript
/*本示例搜索 A1:A3 中的数据。如果发现某单元格的值与上面的一个单元格的值相等，则本示例将显示这个包含重复数据的单元格的地址。*/
function test() {
    let r = Range("A1:A3")
    for (let n = 2; n <= r.Rows.Count; n++) {
        if (r.Cells.Item(n, 1).Value2 == r.Cells.Item(n - 1, 1).Value2) {
            console.log("Duplicate data in " + r.Cells.Item(n, 1).Address())
        }
    }
}
```
