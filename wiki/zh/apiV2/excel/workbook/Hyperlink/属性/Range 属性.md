---
title: "Range (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Hyperlink/%E5%B1%9E%E6%80%A7/Range%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: DAiqdBUS
---

# Range (属性)

返回一个 **[Range](./../../Range/Range 对象.html)** 对象，它代表附加的指定超链接的区域。

## 示例

```javascript
/*此示例显示工作表Sheet1中第二个超链接所在单元格区域的地址*/
function test() {
    console.log(Application.Worksheets.Item("Sheet1").Hyperlinks.Item(2).Range.Address())
}
```

```javascript
/*此示例滚动工作簿窗口，直至超链接区域位于活动窗口的左上角*/
function test() {
    Application.Workbooks.Item(1).Activate()
    let range = Application.ActiveSheet.Hyperlinks.Item(1).Range
    Application.ActiveWindow.ScrollRow = range.Row
    Application.ActiveWindow.ScrollColumn = range.Column
}
```
