---
title: "CompactLayoutRowHeader (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/PivotTable/%E5%B1%9E%E6%80%A7/CompactLayoutRowHeader%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: CGEZ8wfs
---

# CompactLayoutRowHeader (属性)

指定透视数据表在压缩行布局表单中时行标题中显示的标题。可读写 **String**。

## 说明

用户在网格中自定义这些标题时，这些标题也会反映到这些属性中。

## 示例

```javascript
/*本示例将活动工作表中第一张数据透视表压缩行布局显示的行标题设置为 KingSoft*/
function test() {
    ActiveSheet.PivotTables(1).CompactLayoutRowHeader = "KingSoft"
}
```

```javascript
/*本示例显示工作表 Sheet1 中数据透视表1压缩行布局显示的行标题*/
function test() {
    console.log(Worksheets.Item("Sheet1").PivotTables("数据透视表1").CompactLayoutRowHeader)
}
```
