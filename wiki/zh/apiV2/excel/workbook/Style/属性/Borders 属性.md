---
title: "Borders (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Style/%E5%B1%9E%E6%80%A7/Borders%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: BgIX1Mos
---

# Borders (属性)

返回一个 **[Borders](./../../Borders/Borders 对象.html)** 集合，它代表样式或单元格区域（包括定义为条件格式一部分的区域）的边框。

## 示例

```javascript
/*此示例将 Sheet1 中单元格 B2 的底部边框颜色设置为红色细边框。*/
function test() {
    let border = Application.Worksheets.Item("Sheet1").Range("B2").Style.Borders.Item(xlEdgeBottom)
    border.LineStyle = xlContinuous
    border.Weight = xlThin
    border.ColorIndex = 3
}
```

```javascript
/*本示例显示活动工作表中单元格区域左边框粗细是否为薄。*/
function test() {
    let range = ActiveSheet.Range("C3:D5")
    console.log(range.Style.Borders.Item(xlEdgeLeft).Weight == xlThin)
}
```
