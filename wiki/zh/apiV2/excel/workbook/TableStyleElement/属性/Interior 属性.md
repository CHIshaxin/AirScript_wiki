---
title: "Interior (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/TableStyleElement/%E5%B1%9E%E6%80%A7/Interior%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: BNsLiZWe
---

# Interior (属性)

返回一个 **[Interior](./../../Interior/Interior 对象.html)** 对象，该对象代表指定对象的内部。 只读。

## 示例

```javascript
/*本示例将活动工作表上第一张列表的最后一列样式元素的内部设置为红色。*/
function test() {
    let listObj = ActiveSheet.ListObjects.Item(1)
    listObj.TableStyle.TableStyleElements.Item(xlLastColumn).Interior.ColorIndex = 3
}
```

```javascript
/*本示例显示第一张工作表上第一张列表中标题行样式元素内部图案是否为纯色。*/
function test() {
    let listObj = Application.Worksheets.Item(1).ListObjects.Item(1)
    console.log(listObj.TableStyle.TableStyleElements.Item(xlHeaderRow).Interior.Pattern == xlPatternSolid)
}
```
