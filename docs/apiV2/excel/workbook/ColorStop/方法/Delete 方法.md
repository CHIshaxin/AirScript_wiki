---
title: "Delete 方法"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/ColorStop/%E6%96%B9%E6%B3%95/Delete%20%E6%96%B9%E6%B3%95.html
---
# Delete (方法)

删除指定对象。

## 返回值

void

## 示例

```javascript
/*本示例将活动单元格的第一个ColorStop对象删除。*/
function test() {
    let colorStop = ActiveCell.Interior.Gradient.ColorStops.Item(1)
    colorStop.Delete()
}
```

```javascript
/*本示例将当前选定区域的第二个ColorStop对象删除。*/
function test() {
    Range("A1:A10").Select()
    let colorStop = Selection.Interior.Gradient.ColorStops.Item(2)
    colorStop.Delete()
}
```
