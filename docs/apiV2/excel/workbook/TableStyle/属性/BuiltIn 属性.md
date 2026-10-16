---
title: "BuiltIn 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/TableStyle/%E5%B1%9E%E6%80%A7/BuiltIn%20%E5%B1%9E%E6%80%A7.html
---
# BuiltIn (属性)

如果样式为内置样式，则为 **True**。只读 **Boolean** 类型。

## 示例

```javascript
/*本示例显示工作表 Sheet1 上第一张列表的表样式是否为内置样式。*/
function test() {
    let listObj = Application.Worksheets.Item("Sheet1").ListObjects.Item(1)
    console.log(listObj.TableStyle.BuiltIn)
}
```

```javascript
/*本示例显示活动工作簿上的表格样式中的最后一个样式是否为内置样式。*/
function test() {
    let tableStyles = ActiveWorkbook.TableStyles
    console.log(tableStyles.Item(tableStyles.Count).BuiltIn)
}
```
