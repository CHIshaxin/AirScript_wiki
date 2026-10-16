---
title: "Visible 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/LineFormat/%E5%B1%9E%E6%80%A7/Visible%20%E5%B1%9E%E6%80%A7.html
---
# Visible (属性)

返回或设置一个 **MsoTriState** 值，它确定对象是否可见。可读写。

## 示例

```javascript
/*本示例显示活动工作表的第一个形状线条是否可见。*/
function test() {
    let line = ActiveSheet.Shapes.Item(1).Line
    console.log(line.Visible == msoTrue)
}
```

```javascript
/*本示例设置第一张工作表中第二个形状线条不可见。*/
function test() {
    let line = Worksheets.Item(1).Shapes.Item(2).Line
    line.Visible = msoFalse
}
```
