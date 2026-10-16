---
title: "ShowAsAvailableSlicerStyle (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/TableStyle/%E5%B1%9E%E6%80%A7/ShowAsAvailableSlicerStyle%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: CA67Oby1
---

# ShowAsAvailableSlicerStyle (属性)

返回或设置指定的表样式是否作为可用的表样式在切片器样式库中显示。可读/写

## 返回值

Boolean

## 示例

```javascript
/*本示例将活动工作簿上指定的表样式设置为可用的表样式在切片器样式库中显示。*/
function test() {
    let tableStyles = ActiveWorkbook.TableStyles
    tableStyles.Item(286).ShowAsAvailableSlicerStyle = true
}
```

```javascript
/*本示例显示工作表 Sheet1 上第一张列表的表样式是否作为可用的表样式在切片器样式库中显示。*/
function test() {
    let listObj = Application.Worksheets.Item("Sheet1").ListObjects.Item(1)
    console.log(listObj.TableStyle.ShowAsAvailableSlicerStyle)
}
```
