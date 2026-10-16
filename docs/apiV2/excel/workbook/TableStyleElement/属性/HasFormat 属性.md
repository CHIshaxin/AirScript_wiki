---
title: "HasFormat 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/TableStyleElement/%E5%B1%9E%E6%80%A7/HasFormat%20%E5%B1%9E%E6%80%A7.html
---
# HasFormat (属性)

返回表样式元素是否具有应用到指定元素的格式设置。只读 **Boolean** 类型。

## 示例

```javascript
/*本示例显示活动工作表上第一张列表中标题行样式元素是否具有格式设置。*/
function test() {
    let listObj = ActiveSheet.ListObjects.Item(1)
    console.log(listObj.TableStyle.TableStyleElements.Item(xlHeaderRow).HasFormat)
}
```

```javascript
/*本示例显示第一张工作表上第一张列表中总计行样式元素是否具有格式设置。*/
function test() {
    let listObj = Application.Worksheets.Item(1).ListObjects.Item(1)
    console.log(listObj.TableStyle.TableStyleElements.Item(xlGrandTotalRow).HasFormat)
}
```
