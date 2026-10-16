---
title: "Style 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/LineFormat/%E5%B1%9E%E6%80%A7/Style%20%E5%B1%9E%E6%80%A7.html
---
# Style (属性)

返回或设置一个 **MsoLineStyle** 值，它代表线条的样式。

## 示例

```javascript
/*本示例显示活动工作表的第一个形状线条的样式是否为单线。*/
function test() {
    let line = ActiveSheet.Shapes.Item(1).Line
    console.log(line.Style == msoLineSingle)
}
```

```javascript
/*本示例将第一张工作表中第二个形状线条的样式设置为双细线。*/
function test() {
    let line = Worksheets.Item(1).Shapes.Item(2).Line
    line.Style = msoLineThinThin
}
```
