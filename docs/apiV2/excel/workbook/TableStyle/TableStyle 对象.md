---
title: "TableStyle 对象"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/TableStyle/TableStyle%20%E5%AF%B9%E8%B1%A1.html
---
# TableStyle (对象)

代表可应用于表格或切片器的单个样式。

## 说明

表格样式为表格、数据透视表或切片器的一个或所有元素定义格式。例如，列是表格的元素。表格样式可以规定使用交替格式（也称为条带或条纹）对表格中的列进行格式设置。

## 示例

```javascript
/*本示例删除第一张工作簿上的表格样式中的最后一个样式。*/
function test() {
    let tableStyles = Workbooks.Item(1).TableStyles
    tableStyles.Item(tableStyles.Count).Delete()
}
```

```javascript
/*本示例显示第一张工作表上第一张列表的表样式是否为内置样式。*/
function test() {
    let listObj = Application.Worksheets.Item(1).ListObjects.Item(1)
    console.log(listObj.TableStyle.BuiltIn)
}
```
