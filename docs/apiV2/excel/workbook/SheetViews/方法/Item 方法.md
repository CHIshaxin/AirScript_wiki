---
title: "Item 方法"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/SheetViews/%E6%96%B9%E6%B3%95/Item%20%E6%96%B9%E6%B3%95.html
---
# Item (方法)

返回一个 **SheetView** 对象，该对象代表工作簿的视图。只读。

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
| --- | --- | --- | --- |

## 示例

```javascript
/*本示例激活工作表Sheet1，并隐藏第一个工作表视图的标题。*/
function test() {
    Application.Worksheets.Item("Sheet1").Activate()
    ActiveWindow.SheetViews.Item(1).DisplayHeadings = false
}
```

```javascript
/*本示例更改第二个工作表视图所在的工作表名称。*/
function test() {
    let sheetView = ActiveWindow.SheetViews.Item(2)
    sheetView.Sheet.Name = "成绩表"
}
```
