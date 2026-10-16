---
title: "TableStyleElement 对象"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/TableStyleElement/TableStyleElement%20%E5%AF%B9%E8%B1%A1.html
---
# TableStyleElement (对象)

代表单个表格样式元素。

## 说明

表格样式为表格、数据透视表或切片器的一个或所有元素定义格式。例如，标题行是表格的元素。表格样式可以规定标题行的填充色为红色。

表格中每个表格样式元素的格式设置可在适用于该元素的表格样式中指定。

## 示例

```javascript
/*本示例清除活动工作表上第一张列表中总计行样式元素的格式。*/
function test() {
    let listObj = ActiveSheet.ListObjects.Item(1)
    listObj.TableStyle.TableStyleElements.Item(xlGrandTotalRow).Clear()
}
```

```javascript
/*本示例显示第一张工作表上第一张列表的行条纹1样式元素条带的大小。*/
function test() {
    let listObj = Worksheets.Item(1).ListObjects.Item(1)
    console.log(listObj.TableStyle.TableStyleElements.Item(xlRowStripe1).StripeSize)
}
```
