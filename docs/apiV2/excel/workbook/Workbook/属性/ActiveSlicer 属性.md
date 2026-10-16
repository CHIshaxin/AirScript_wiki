---
title: "ActiveSlicer 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Workbook/%E5%B1%9E%E6%80%A7/ActiveSlicer%20%E5%B1%9E%E6%80%A7.html
---
# ActiveSlicer (属性)

返回一个对象，该对象代表活动工作簿或指定工作簿中的活动切片器。如果没有切片器处于活动状态，则返回 **null**。只读。

## 说明

如果没有指定 **[Workbook](./../../Workbook/Workbook 对象.html)** 对象，则此属性将返回活动工作簿中的活动切片器。

## 返回值

Slicer

## 示例

```javascript
/*本示例设置活动工作簿上活动切片器的标题。*/
function test() {
    ActiveWorkbook.ActiveSlicer.Caption = "nameid"
}
```

```javascript
/*本示例设置第一张工作簿上活动切片器内的列数。*/
function test() {
    Workbooks.Item(1).ActiveSlicer.NumberOfColumns = 2
}
```
