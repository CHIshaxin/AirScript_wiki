---
title: "TableStyleElements 对象"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/TableStyleElements/TableStyleElements%20%E5%AF%B9%E8%B1%A1.html
---
# TableStyleElements (对象)

代表表格样式元素。

## 说明

表格样式为表格、数据透视表或切片器的一个或所有元素定义格式。例如，标题行、最后一列或总计行是表格的元素，表格样式可以规定标题行的填充色为蓝色，最后一列为红色。

表格中表格样式元素的格式设置可在适用于该元素的表格样式中指定。 **[XlTableStyleElementType](./../枚举/XlTableStyleElementType 枚举.html)** 枚举包含可供使用的表格样式元素的类型。

## 示例

```javascript
/*本示例显示第一张工作表上第一张列表是否存在表格样式元素。*/
function test() {
    let listObj = Application.Worksheets.Item(1).ListObjects.Item(1)
    console.log(listObj.TableStyle.TableStyleElements.Count > 0)
}
```

```javascript
/*本示例将活动工作表上第一张列表的最后一列样式元素的内部设置为红色。*/
function test() {
    let listObj = ActiveSheet.ListObjects.Item(1)
    listObj.TableStyle.TableStyleElements.Item(xlLastColumn).Interior.ColorIndex = 3
}
```
