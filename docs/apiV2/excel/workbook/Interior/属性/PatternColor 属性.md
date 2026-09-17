---
title: "PatternColor 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Interior/%E5%B1%9E%E6%80%A7/PatternColor%20%E5%B1%9E%E6%80%A7.html
---
# PatternColor (属性)

以 RGB 值返回或设置内部图案的颜色。可读/写 **Variant** 类型。

## 示例

```javascript
/*本示例对 Sheet1 中 A1:C3 区域单元格添加棋盘图案，并将图案设置为红色。*/
function test() {
    let interior = Application.Worksheets.Item("Sheet1").Range("A1:C3").Interior
    interior.Pattern = xlPatternChecker
    interior.PatternColor = RGB(255, 0, 0)
}
```

```javascript
/*本示例判断活动工作表 A1:C5 区域图案颜色是否为蓝色。*/
function test() {
    let interior = ActiveSheet.Range("A1:C5").Interior
    console.log(interior.PatternColor == RGB(0, 0, 255))
}
```
