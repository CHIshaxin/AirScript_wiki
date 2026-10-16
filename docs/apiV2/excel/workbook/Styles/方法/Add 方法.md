---
title: "Add 方法"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Styles/%E6%96%B9%E6%B3%95/Add%20%E6%96%B9%E6%B3%95.html
---
# Add (方法)

新建样式并将其添加到当前工作簿的可用样式列表中。返回一个代表新样式的 **[Style](./../../Style/Style 对象.html)** 对象。

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| Name | string | 必填 | 新样式的名称。 |

## 示例

```javascript
/*本示例定义新样式，并将工作表 Sheet1 中的单元格 A1 改为此样式。*/
function test() {
    let Sty = Application.ActiveWorkbook.Styles.Add("theNewStyle")
    Sty.IncludeNumber = false
    Sty.IncludeFont = true
    Sty.IncludeAlignment = false
    Sty.IncludeBorder = false
    Sty.IncludePatterns = false
    Sty.IncludeProtection = false
    Sty.Font.Name = "Arial"
    Sty.Font.Size = 18
    Application.Worksheets.Item("Sheet1").Range("A1").Style = "theNewStyle"
}
```

```javascript
/*此示例为活动工作簿添加一个新的样式，并设置该样式内部填充颜色为绿色，再将单元格区域A5:A10改用此样式。*/
function test() {
    let sty = Application.ActiveWorkbook.Styles.Add("新样式")
    sty.Interior.Color = RGB(0, 255, 0)
    Application.Worksheets.Item(1).Range("A5:A10").Style = "新样式"
}
```
