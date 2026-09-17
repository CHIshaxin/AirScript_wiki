---
title: "NumberOfColumns 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Slicer/%E5%B1%9E%E6%80%A7/NumberOfColumns%20%E5%B1%9E%E6%80%A7.html
---
# NumberOfColumns (属性)

返回或设置指定切片器内的列数。可读/写。

## 说明

默认值为 1。

## 返回值

Long

## 示例

```javascript
/*本示例显示工作表 Sheet1 上第一张数据透视表的第二个切片器内的列数。*/
function test() {
    console.log(Worksheets.Item("Sheet1").PivotTables(1).Slicers.Item(2).NumberOfColumns)
}
```

```javascript
/*本示例将活动工作表上数据透视表的第一个切片器内的列数设置为 3。*/
function test() {
    let slicer = ActiveSheet.Range("I1").PivotTable.Slicers.Item(1)
    slicer.NumberOfColumns = 3
}
```
