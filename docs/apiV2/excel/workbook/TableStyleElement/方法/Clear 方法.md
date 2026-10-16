---
title: "Clear 方法"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/TableStyleElement/%E6%96%B9%E6%B3%95/Clear%20%E6%96%B9%E6%B3%95.html
---
# Clear (方法)

清除此元素的格式。

## 示例

```javascript
/*本示例清除活动工作表上第一张列表的样式元素的格式。*/
function test() {
    let listObj = ActiveSheet.ListObjects.Item(1)
    listObj.TableStyle.TableStyleElements.Item(xlWholeTable).Clear()
}
```

```javascript
/*本示例清除第一张工作表上第一张列表中标题行样式元素的格式。*/
function test() {
    let listObj = Application.Worksheets.Item(1).ListObjects.Item(1)
    let tableStyleElement = listObj.TableStyle.TableStyleElements.Item(xlHeaderRow)
    tableStyleElement.Clear()
}
```
