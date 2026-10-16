---
title: "Item 方法"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/TableStyleElements/%E6%96%B9%E6%B3%95/Item%20%E6%96%B9%E6%B3%95.html
---
# Item (方法)

从集合中返回一个对象。

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
| --- | --- | --- | --- |

## 返回值

TableStyleElement

## 示例

```javascript
/*本示例将第一张工作表上第一张列表的表样式中第一列的字体加粗。*/
function test() {
    let listObj = Application.Worksheets.Item(1).ListObjects.Item(1)
    let tableStyleElement = listObj.TableStyle.TableStyleElements.Item(xlFirstColumn)
    tableStyleElement.Font.Bold = true
}
```

```javascript
/*本示例显示活动工作表上第一张列表的表样式的标题行是否具有格式设置。*/
function test() {
    let listObj = ActiveSheet.ListObjects.Item(1)
    console.log(listObj.TableStyle.TableStyleElements.Item(xlHeaderRow).HasFormat)
}
```
