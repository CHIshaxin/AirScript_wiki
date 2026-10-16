---
title: "Name (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/TableStyle/%E5%B1%9E%E6%80%A7/Name%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: i_ZxUJIq
---

# Name (属性)

返回对象的名称。只读 **String** 类型。

## 示例

```javascript
/*本示例显示活动工作表上包含 C3 单元格的列表的表样式的名称。*/
function test() {
    let listObj = Application.ActiveSheet.Range("C3").ListObject
    console.log(listObj.TableStyle.Name)
}
```

```javascript
/*本示例将第一张工作簿的表格样式中第一个样式名称赋值给D2单元格。*/
function test() {
    Range("D2").Value2 = Workbooks.Item(1).TableStyles.Item(1).Name
}
```
