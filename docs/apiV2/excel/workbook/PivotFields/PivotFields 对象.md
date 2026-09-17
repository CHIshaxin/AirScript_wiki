---
title: "PivotFields 对象"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/PivotFields/PivotFields%20%E5%AF%B9%E8%B1%A1.html
---
# PivotFields (对象)

指定的数据透视表中所有 **[PivotField](./../PivotField/PivotField 对象.html)** 对象的集合。

## 说明

下列存取器方法返回数据透视表字段的子集，在某些情况下，使用这些属性会更为方便：

- **ColumnFields** 属性
- **DataFields** 属性
- **HiddenFields** 属性
- **PageFields** 属性
- **RowFields** 属性
- **VisibleFields** 属性

使用 **[PivotTable](./../PivotTable/PivotTable 对象.html)** 对象的 **PivotFields** 方法可返回 **[PivotFields](./../PivotFields/PivotFields 对象.html)** 集合。

## 示例

```javascript
/*本示例显示 Sheet1 上第一张数据透视表中的字段名称。*/
function test() {
    let pTab = Worksheets.Item("Sheet1").PivotTables(1)
    for (let i = 1; i <= pTab.PivotFields().Count; i++) {
        console.log(pTab.PivotFields(i).Name)
    }
}
```

使用 **PivotFields** (*index*)（其中 *index* 是字段名称或索引号）可返回一个 **[PivotField](./../PivotField/PivotField 对象.html)** 对象。

```javascript
/*本示例使 Sheet1 上第一张数据透视表中的字段“name”成为行字段。*/
function test() {
    Worksheets.Item("sheet1").PivotTables(1).PivotFields("name").Orientation = xlRowField
}
```
