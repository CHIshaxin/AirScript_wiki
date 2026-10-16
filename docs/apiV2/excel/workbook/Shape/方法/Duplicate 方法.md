---
title: "Duplicate 方法"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Shape/%E6%96%B9%E6%B3%95/Duplicate%20%E6%96%B9%E6%B3%95.html
---
# Duplicate (方法)

复制对象，并返回对新复制对象的引用。

## 返回值

Shape

## 示例

```javascript
/*此示例复制工作表 Sheet1 上第一个形状，然后选择新复制的形状。*/
function test() {
    let duplicate = Worksheets("Sheet1").Shapes.Item(1).Duplicate()
    duplicate.Select()
}
```

```javascript
/*此示例复制活动工作表上第二个形状并调整大小。*/
function test() {
    let duplicate = ActiveSheet.Shapes.Item(2).Duplicate()
    duplicate.Width *= 2
    duplicate.Height *= 2
}
```
