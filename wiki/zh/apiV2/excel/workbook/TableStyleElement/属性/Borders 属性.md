---
title: "Borders (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/TableStyleElement/%E5%B1%9E%E6%80%A7/Borders%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: YFL_cos-
---

# Borders (属性)

返回一个代表表样式元素的边框的 **[Borders](./../../Borders/Borders 对象.html)** 集合。只读。

## 示例

```javascript
/*本示例将表格的上边框设置为红色实线的细边框。*/
function test() {
    let border = Application.ActiveWorkbook.TableStyles.Item("Table Style 4").TableStyleElements.Item(xlWholeTable).Borders.Item(xlEdgeTop)
    border.Color = RGB(255, 0, 0)
    border.TintAndShade = 0
    border.Weight = xlThin
    border.LineStyle = xlContinuous
}
```

```javascript
/*本示例显示活动工作表上第一张列表中标题行样式元素的边框主题颜色是否为蓝色。*/
function test() {
    let listObj = ActiveSheet.ListObjects.Item(1)
    console.log(listObj.TableStyle.TableStyleElements.Item(xlHeaderRow).Borders.ThemeColor == RGB(0, 0, 255))
}
```
