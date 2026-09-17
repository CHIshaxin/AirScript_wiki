---
title: "ShowHeaders 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/ListObject/%E5%B1%9E%E6%80%A7/ShowHeaders%20%E5%B1%9E%E6%80%A7.html
---
# ShowHeaders (属性)

返回或设置是否显示指定的 **[ListObject](./../../ListObject/ListObject 对象.html)** 对象的标题信息。可读/写 **Boolean** 类型。

## 示例

```javascript
/*本示例显示工作表 Sheet1 上第一张列表是否展示标题信息。*/
function test() {
    let listObj = Worksheets.Item("Sheet1").ListObjects.Item(1)
    console.log(listObj.ShowHeaders)
}
```

```javascript
/*本示例将活动工作表上包含 A2 单元格的列表设置为不展示标题信息。*/
function test() {
    let listObj = Application.ActiveSheet.Range("A2").ListObject
    listObj.ShowHeaders = false
}
```
