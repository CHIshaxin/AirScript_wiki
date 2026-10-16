---
title: "TableStyleElements (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/TableStyle/%E5%B1%9E%E6%80%A7/TableStyleElements%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: DWJaUPXo
---

# TableStyleElements (属性)

返回 **[TableStyleElements](./../../TableStyleElements/TableStyleElements 对象.html)** 对象，只读。

## 示例

```javascript
/*本示例显示第一张工作表上第一张列表的表样式元素的数量。*/
function test() {
    let listObj = Application.Worksheets.Item(1).ListObjects.Item(1)
    console.log(listObj.TableStyle.TableStyleElements.Count)
}
```

```javascript
/*本示例将活动工作簿上所有表样式中第一列的字体加粗。*/
function test() {
    let tableStyles = ActiveWorkbook.TableStyles
    for (let i = 1; i <= tableStyles.Count; i++) {
        tableStyles.Item(i).TableStyleElements.Item(xlFirstColumn).Font.Bold = true
    }
}
```
