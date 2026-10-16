---
title: "Range 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/PivotCell/%E5%B1%9E%E6%80%A7/Range%20%E5%B1%9E%E6%80%A7.html
---
# Range (属性)

返回一个 **[Range](./../../Range/Range 对象.html)** 对象，它代表指定的 PivotCell 所应用的区域。

## 示例

```javascript
/*本示例将包含“Sheet1”工作表单元格的数据透视项的地址保存在变量中并显示。*/
function test() {
    let address = Worksheets.Item("Sheet1").Range("I5").PivotCell.Range.Address()
    console.log(address)
}
```

```javascript
/*此示例滚动工作表窗口，直至包含活动单元格的数据透视项区域位于活动窗口的左上角。*/
function test() {
    Worksheets.Item(1).Activate()
    let range = ActiveCell.PivotCell.Range
    ActiveWindow.ScrollRow = range.Row
    ActiveWindow.ScrollColumn = range.Column
}
```
