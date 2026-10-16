---
title: "Style (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Slicer/%E5%B1%9E%E6%80%A7/Style%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: Cx2qnhtE
---

# Style (属性)

返回或设置当前应用于指定切片器的样式。可读/写。

## 返回值

Variant

## 示例

```javascript
/*本示例显示活动工作表上第一张数据透视表的第三个切片器的样式名称。*/
function test() {
    console.log(ActiveSheet.PivotTables(1).Slicers.Item(3).Style.Name)
}
```

```javascript
/*本示例设置工作表 Sheet1 上第一张数据透视表的第一个切片器的样式。*/
function test() {
    Worksheets.Item("Sheet1").PivotTables(1).Slicers.Item(1).Style = "SlicerStyleLight3"
}
```
