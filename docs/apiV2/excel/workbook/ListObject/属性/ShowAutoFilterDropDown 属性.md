---
title: "ShowAutoFilterDropDown 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/ListObject/%E5%B1%9E%E6%80%A7/ShowAutoFilterDropDown%20%E5%B1%9E%E6%80%A7.html
---
# ShowAutoFilterDropDown (属性)

显示 **[ListObject](./../../ListObject/ListObject 对象.html)** 对象的“自动筛选”下拉列表时为 **True**。 可读写 ，**Boolean** 类型。

## 示例

```javascript
/*本示例显示工作表 Sheet1 上第一张列表是否展示“自动筛选”下拉列表。*/
function test() {
    let listObj = Worksheets.Item("Sheet1").ListObjects.Item(1)
    console.log(listObj.ShowAutoFilterDropDown)
}
```

```javascript
/*本示例判断如果活动工作表上包含 C2 单元格的列表的名称为“表1”，则将该列表设置为展示“自动筛选”下拉列表。*/
function test() {
    let listObj = Application.ActiveSheet.Range("C2").ListObject
    if (listObj.Name == "表1") {
        listObj.ShowAutoFilterDropDown = true
    }
}
```
