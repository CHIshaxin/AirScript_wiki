---
title: "Size (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Font/%E5%B1%9E%E6%80%A7/Size%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: BG4WKKOQ
---

# Size (属性)

返回或设置字号。可读/写 **Variant** 类型。

## 示例

```javascript
/*本示例将 Sheet1 的 A1:D10 单元格的字体大小设为 12 磅*/
function test() {
    let range = Application.Worksheets.Item("Sheet1").Range("A1:D10")
    range.Value2 = "Test"
    range.Font.Size = 12
}
```

```javascript
/*本示例显示第一个工作表中单元格E1的字体大小*/
function test() {
    let font = Application.Worksheets.Item(1).Range("E1").Font
    console.log(`字体大小为 ${font.Size}`)
}
```

``
