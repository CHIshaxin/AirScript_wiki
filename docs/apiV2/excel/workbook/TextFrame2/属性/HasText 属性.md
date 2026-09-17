---
title: "HasText 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/TextFrame2/%E5%B1%9E%E6%80%A7/HasText%20%E5%B1%9E%E6%80%A7.html
---
# HasText (属性)

返回指定文本框是否有文本。只读 **MsoTriState** 类型。

## 示例

```javascript
/*本示例如果文本框包含文本，则设置文本框中的文本字体为“Arial”。*/
function test() {
    let textFrame2 = Application.ActiveSheet.Shapes.Item(1).TextFrame2
    if (textFrame2.HasText) {
        textFrame2.TextRange.Font.Name = "Arial"
    }
}
```

```javascript
/*本示例显示第一张工作表中第二个形状的文本框是否包含文本。*/
function test() {
    let textFrame2 = Application.Worksheets.Item(1).Shapes.Item(2).TextFrame2
    console.log(textFrame2.HasText == msoTrue)
}
```
