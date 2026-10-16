---
title: "Pattern (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/LineFormat/%E5%B1%9E%E6%80%A7/Pattern%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: IFdPDsfR
---

# Pattern (属性)

返回或设置一个代表填充图案的 **MsoPatternType** 值。

## 示例

```javascript
/*本示例将第一张工作表中第二个形状线条的图案设置为十字形。*/
function test() {
    let line = Worksheets.Item(1).Shapes.Item(2).Line
    line.Pattern = msoPatternCross
}
```

```javascript
/*本示例显示活动工作表的第一个形状线条图案是否为混和图案。*/
function test() {
    let line = ActiveSheet.Shapes.Item(1).Line
    console.log(line.Pattern == msoPatternMixed)
}
```
