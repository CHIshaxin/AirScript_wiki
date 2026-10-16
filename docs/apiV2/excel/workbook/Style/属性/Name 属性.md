---
title: "Name 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Style/%E5%B1%9E%E6%80%A7/Name%20%E5%B1%9E%E6%80%A7.html
---
# Name (属性)

返回一个 **String** 值，它代表对象的名称。

## 示例

```javascript
/*首先用宏语言，然后用用户语言显示活动工作簿中的第一种样式的名称。*/
function test() {
    let sty = Application.ActiveWorkbook.Styles.Item(1)
    console.log("The name of the style: " + sty.Name)
    console.log("The localized name of the style: " + sty.NameLocal)
}
```

```javascript
/*显示活动工作簿的 Sheet1 中的默认 ListObject 对象的名称。*/
function test() {
    let wrksht = Application.ActiveWorkbook.Worksheets.Item("Sheet1")
    let oListObj = wrksht.ListObjects.Item(1)
    console.log(oListObj.Name)
}
```
