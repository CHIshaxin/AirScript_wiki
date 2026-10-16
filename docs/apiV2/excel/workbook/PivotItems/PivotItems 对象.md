---
title: "PivotItems 对象"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/PivotItems/PivotItems%20%E5%AF%B9%E8%B1%A1.html
---
# PivotItems (对象)

数据透视表字段中所有 **[PivotItem](./../PivotItem/PivotItem 对象.html)** 对象的集合。

## 说明

这些项目是某个字段类型中的各个数据项。

使用 **PivotItems** 方法可返回 **[PivotItems](./../PivotItems/PivotItems 对象.html)** 集合。

## 示例

```javascript
/*本示例为 Sheet1 上第一张数据透视表创建那些字段中包含的字段名和项目的枚举列表。*/
function test() {
    Worksheets.Item("Sheet2").Activate()
    let pTab = Worksheets.Item("Sheet1").PivotTables(1)
    let c = 1
    for (let i = 1; i <= pTab.PivotFields().Count; i++) {
        let r = 1
        Cells.Item(r, c).Value2 = pTab.PivotFields(i).Name
        r++
        for (let x = 1; x <= pTab.PivotFields(i).PivotItems().Count; x++) {
            Cells.Item(r, c).Value2 = pTab.PivotFields(i).PivotItems(x).Name
            r++
        }
        c++
    }
}
```

使用 **PivotItems** (*index*)（其中 *index* 是项目索引号或名称）可返回一个 **[PivotItem](./../PivotItem/PivotItem 对象.html)** 对象。

```javascript
/*此示例隐藏工作表 Sheet1 中第一张数据透视表的字段“id”的数据项“4”。*/
function test() {
    let pvtItems = Worksheets.Item("sheet1").PivotTables(1).PivotFields("id").PivotItems()
    pvtItems.Item("4").Visible = false
}
```
