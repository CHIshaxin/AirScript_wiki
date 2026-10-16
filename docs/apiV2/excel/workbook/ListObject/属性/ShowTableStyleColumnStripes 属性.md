---
title: "ShowTableStyleColumnStripes 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/ListObject/%E5%B1%9E%E6%80%A7/ShowTableStyleColumnStripes%20%E5%B1%9E%E6%80%A7.html
---
# ShowTableStyleColumnStripes (属性)

返回或设置指定的 **[ListObject](./../../ListObject/ListObject 对象.html)** 对象是否使用“列条纹”表样式。可读/写 **Boolean** 类型。

## 示例

```javascript
/*本示例显示活动工作表上包含 C2 单元格的列表是否使用“列条纹”表样式。*/
function test() {
    let listObj = Application.ActiveSheet.Range("C2").ListObject
    console.log(listObj.ShowTableStyleColumnStripes)
}
```

```javascript
/*本示例将工作表 Sheet1 上第一张列表设置为使用“列条纹”表样式。*/
function test() {
    let listObj = Worksheets.Item("Sheet1").ListObjects.Item(1)
    listObj.ShowTableStyleColumnStripes = true
}
```
