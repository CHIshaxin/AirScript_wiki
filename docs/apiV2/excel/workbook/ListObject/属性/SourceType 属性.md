---
title: "SourceType 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/ListObject/%E5%B1%9E%E6%80%A7/SourceType%20%E5%B1%9E%E6%80%A7.html
---
# SourceType (属性)

返回一个 **[XlListObjectSourceType](./../../枚举/XlListObjectSourceType 枚举.html)** 值，它代表列表的当前源。

## 示例

```javascript
/*本示例显示 XlListObjectSourceType 常量，该常量表示活动工作簿的 Sheet1 上的默认列表的源。*/
function test() {
    let listObj = Application.ActiveWorkbook.Worksheets.Item("Sheet1").ListObjects.Item(1)
    console.log(listObj.SourceType)
}
```

```javascript
/*本示例显示活动工作表上包含 C3 单元格的列表的 SourceType 属性是否为 xlSrcRange。*/
function test() {
    let listObj = Application.ActiveSheet.Range("C3").ListObject
    console.log(listObj.SourceType == xlSrcRange)
}
```
