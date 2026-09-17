---
title: "Font 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/TableStyleElement/%E5%B1%9E%E6%80%A7/Font%20%E5%B1%9E%E6%80%A7.html
---
# Font (属性)

返回一个 **[Font](./../../Font/Font 对象.html)** 对象，它代表指定对象的字体。 只读。

## 示例

```javascript
/*本示例将第一张工作表上第一张列表中第一列样式元素的字体加粗。*/
function test() {
    let listObj = Application.Worksheets.Item(1).ListObjects.Item(1)
    let tableStyleElement = listObj.TableStyle.TableStyleElements.Item(xlFirstColumn)
    tableStyleElement.Font.Bold = true
}
```

```javascript
/*本示例显示活动工作表上第一张列表的样式元素的字体样式是否为斜体。*/
function test() {
    let listObj = ActiveSheet.ListObjects.Item(1)
    console.log(listObj.TableStyle.TableStyleElements.Item(xlWholeTable).Font.FontStyle == "Italic")
}
```
