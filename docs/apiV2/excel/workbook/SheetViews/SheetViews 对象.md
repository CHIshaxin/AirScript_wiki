---
title: "SheetViews 对象"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/SheetViews/SheetViews%20%E5%AF%B9%E8%B1%A1.html
---
# SheetViews (对象)

指定的或活动工作簿窗口中所有工作表视图的集合。

## 说明

## 示例

```javascript
/*本示例显示活动窗口的工作表视图数量*/
function test() {
    let count = Application.ActiveWindow.SheetViews.Count
    console.log(`活动窗口的工作表视图数量是 ${count}`)
}
```

```javascript
/*本示例更改第一个工作表视图所在的工作表名称，并显示出来。*/
function test() {
    let sheetview = ActiveWindow.SheetViews.Item(1)
    sheetview.Sheet.Name = "成绩表"
    console.log(sheetview.Sheet.Name)
}
```
